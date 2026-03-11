"""
1D Row-Block Subset + Placement Model using OR-Tools CP-SAT directly.

Replaces the former cpmpy-based implementation so the code runs on any
ortools version (including 9.15+) without cpmpy compatibility issues.

Public interface is identical to the previous version:
  model = RowBlock1DOrderModel(lengths_cm, heights_cm, weights_kg, values, ...)
  solved = model.solve(solver='ortools', time_limit=5)
  order  = model.loaded_indices_in_order()   # list of 1-based block indices
  model.usedLen.value()
  model.loadedValue.value()
  model.loadedWeight.value()
"""

from ortools.sat.python import cp_model

from config import (
    CONTAINER_LENGTH_CM, CONTAINER_DOOR_HEIGHT_CM,
    CONTAINER_MAX_WEIGHT_KG, ROW_GAP_CM,
)


class _ValueProxy:
    """Mimics cpmpy's intvar.value() so pipeline.py needs no changes."""
    def __init__(self, solver, var):
        self._solver = solver
        self._var = var

    def value(self):
        return self._solver.Value(self._var)


class RowBlock1DOrderModel:
    """
    1D Row-Block Subset + Placement Model (single container).

    Decision: slot[r] in {0..N}
      slot[r] = i  → block i is placed at row position r (back → door)
      slot[r] = 0  → row r is empty

    Hard constraints:
      * length:  sum(row lengths) + gap*(rowsUsed-1) <= L
      * weight:  sum(row weights) <= Wmax
      * heights: non-increasing back → door
      * door:    last used row height <= Hdoor

    Objective: maximise BIG*loadedValue + usedLen  (lex-like)
    """

    def __init__(
        self,
        lengths_cm, heights_cm, weights_kg, values,
        L_cm=CONTAINER_LENGTH_CM,
        gap_cm=ROW_GAP_CM,
        Wmax_kg=CONTAINER_MAX_WEIGHT_KG,
        Hdoor_cm=CONTAINER_DOOR_HEIGHT_CM,
        Rmax=None,
        unload_limit=None,
        min_loaded_value=None,
    ):
        self.L     = int(L_cm)
        self.g     = int(gap_cm)
        self.Wmax  = int(Wmax_kg)
        self.Hdoor = int(Hdoor_cm)

        self.len_in = [int(x) for x in lengths_cm]
        self.h_in   = [int(x) for x in heights_cm]
        self.w_in   = [int(x) for x in weights_kg]
        self.val_in = [int(x) for x in values]

        self.N = len(self.len_in)
        assert self.N == len(self.h_in) == len(self.w_in) == len(self.val_in)

        # 0-padded lookup arrays (index 0 = empty slot, contributes 0)
        self.len0 = [0] + self.len_in
        self.h0   = [0] + self.h_in
        self.w0   = [0] + self.w_in
        self.val0 = [0] + self.val_in

        if Rmax is None:
            min_len = min(self.len_in) if self.N > 0 else self.L
            Rmax = (self.L + self.g) // (min_len + self.g) if (min_len + self.g) > 0 else self.N
            Rmax = max(1, min(Rmax, self.N))
        self.Rmax = int(Rmax)

        self._cp_model  = cp_model.CpModel()
        self._cp_solver = cp_model.CpSolver()

        self._build(unload_limit, min_loaded_value)

    # ------------------------------------------------------------------
    def _build(self, unload_limit, min_loaded_value):
        m = self._cp_model
        N = self.N
        R = self.Rmax
        g = self.g

        # ── Variables ──────────────────────────────────────────────────
        slot = [m.NewIntVar(0, N, f'slot_{r}') for r in range(R)]
        used = [m.NewBoolVar(f'used_{r}')       for r in range(R)]

        rows_used  = m.NewIntVar(0, R,          'rowsUsed')
        gap_count  = m.NewIntVar(0, max(R-1,0), 'gapCount')
        used_len   = m.NewIntVar(0, self.L,     'usedLen')
        loaded_wt  = m.NewIntVar(0, self.Wmax,  'loadedWeight')
        total_val  = sum(self.val_in) if self.val_in else 0
        loaded_val = m.NewIntVar(0, total_val,  'loadedValue')

        max_len = max(self.len0)
        max_h   = max(self.h0)
        max_w   = max(self.w0)
        max_val = max(self.val0) if self.val0 else 0

        # Element proxy vars: X_slot[r] = X0[slot[r]]
        len_slot = [m.NewIntVar(0, max_len, f'ls_{r}') for r in range(R)]
        h_slot   = [m.NewIntVar(0, max_h,   f'hs_{r}') for r in range(R)]
        w_slot   = [m.NewIntVar(0, max_w,   f'ws_{r}') for r in range(R)]
        val_slot = [m.NewIntVar(0, max_val, f'vs_{r}') for r in range(R)]

        # Store refs needed by pipeline helpers
        self._slot       = slot
        self._used_len   = used_len
        self._loaded_val = loaded_val
        self._loaded_wt  = loaded_wt

        # ── Element lookups ────────────────────────────────────────────
        for r in range(R):
            m.AddElement(slot[r], self.len0, len_slot[r])
            m.AddElement(slot[r], self.h0,   h_slot[r])
            m.AddElement(slot[r], self.w0,   w_slot[r])
            m.AddElement(slot[r], self.val0, val_slot[r])

        # ── C0: used[r] <-> slot[r] != 0 ──────────────────────────────
        for r in range(R):
            m.Add(slot[r] >= 1).OnlyEnforceIf(used[r])
            m.Add(slot[r] == 0).OnlyEnforceIf(used[r].Not())

        # ── C1: contiguity — once empty, always empty ──────────────────
        for r in range(R - 1):
            m.AddImplication(used[r].Not(), used[r + 1].Not())

        # ── C2: no duplicate non-zero slots ───────────────────────────
        # Encode via auxiliary vars: aux[r] = slot[r] if used, else N+1+r
        # Then AllDifferent(aux) guarantees unique block assignments.
        aux_slots = []
        for r in range(R):
            aux = m.NewIntVar(1, N + R, f'aux_{r}')
            m.Add(aux == slot[r]).OnlyEnforceIf(used[r])
            m.Add(aux == N + 1 + r).OnlyEnforceIf(used[r].Not())
            aux_slots.append(aux)
        m.AddAllDifferent(aux_slots)

        # ── C3: rowsUsed ───────────────────────────────────────────────
        m.Add(rows_used == sum(used))

        # ── C4: gapCount = max(0, rowsUsed - 1) ───────────────────────
        any_used = m.NewBoolVar('any_used')
        m.Add(rows_used >= 1).OnlyEnforceIf(any_used)
        m.Add(rows_used == 0).OnlyEnforceIf(any_used.Not())
        m.Add(gap_count == 0).OnlyEnforceIf(any_used.Not())
        m.Add(gap_count == rows_used - 1).OnlyEnforceIf(any_used)

        # ── C5: length constraint ──────────────────────────────────────
        m.Add(used_len == sum(len_slot) + g * gap_count)
        m.Add(used_len <= self.L)

        # ── C6: weight constraint ──────────────────────────────────────
        m.Add(loaded_wt == sum(w_slot))
        m.Add(loaded_wt <= self.Wmax)

        # ── C7: value ─────────────────────────────────────────────────
        m.Add(loaded_val == sum(val_slot))

        # ── C8: height ordering — non-increasing back → door ──────────
        for r in range(R - 1):
            m.Add(h_slot[r] >= h_slot[r + 1])

        # ── C9: door height on the last used row ──────────────────────
        for r in range(R):
            if r < R - 1:
                m.Add(h_slot[r] <= self.Hdoor).OnlyEnforceIf(
                    [used[r], used[r + 1].Not()]
                )
            else:
                m.Add(h_slot[r] <= self.Hdoor).OnlyEnforceIf(used[r])

        # ── Optional constraints ───────────────────────────────────────
        if unload_limit is not None:
            m.Add(N - rows_used <= int(unload_limit))
        if min_loaded_value is not None:
            m.Add(loaded_val >= int(min_loaded_value))

        # ── Objective ─────────────────────────────────────────────────
        BIG = 10 ** 6
        m.Maximize(BIG * loaded_val + used_len)

    # ------------------------------------------------------------------
    def solve(self, solver='ortools', time_limit=None):
        if time_limit is not None:
            self._cp_solver.parameters.max_time_in_seconds = float(time_limit)
        # Support both old (Solve) and new (solve) ortools CpSolver API
        if hasattr(self._cp_solver, 'Solve'):
            status = self._cp_solver.Solve(self._cp_model)
        else:
            status = self._cp_solver.solve(self._cp_model)
        return status in (cp_model.OPTIMAL, cp_model.FEASIBLE)

    # ------------------------------------------------------------------
    # Pipeline-compatible interface (mirrors old cpmpy-based class)
    # ------------------------------------------------------------------
    @property
    def usedLen(self):
        return _ValueProxy(self._cp_solver, self._used_len)

    @property
    def loadedValue(self):
        return _ValueProxy(self._cp_solver, self._loaded_val)

    @property
    def loadedWeight(self):
        return _ValueProxy(self._cp_solver, self._loaded_wt)

    def loaded_indices_in_order(self):
        """Return block indices (1..N) in back→door order, excluding 0."""
        return [
            int(self._cp_solver.Value(self._slot[r]))
            for r in range(self.Rmax)
            if int(self._cp_solver.Value(self._slot[r])) != 0
        ]

    def unloaded_indices(self):
        loaded = set(self.loaded_indices_in_order())
        return [i for i in range(1, self.N + 1) if i not in loaded]

    def compute_y_starts(self):
        order = self.loaded_indices_in_order()
        y = 0
        out = []
        for idx in order:
            out.append((idx, y))
            y += self.len0[idx] + self.g
        return out
