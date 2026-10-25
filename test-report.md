# Test evidence template

Environment: OS / Python / Excel build / Office bitness / locale / references.
Fixture: source hash, authorized working-copy identity, expected-result authority.

| Case | Layer | Preconditions | Action | Oracle | Observed | Status | Evidence location |
|---|---|---|---|---|---|---|---|

Layers: STATIC_REVIEW, PYTHON_TEST, MOCK_COM, VBA_COMPILED, VBA_RUNTIME, REAL_COM,
UI, END_TO_END. Status: PASS, FAIL, NOT_RUN, BLOCKED.

Record application state before/after, owned resources, cleanup failures, retained
artifacts and limitations. Do not infer process exit from Quit returning. Do not
mark user acceptance or live agent routing passed from static file assertions.
