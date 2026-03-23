# Testing

!!! warning "Coming in next sprint"
The testing infrastructure is being finalised. This page will document the full testing strategy once the test suite is in place.

    **Planned content:**

    - Unit test approach: `Protocol`-based fakes that stub COM objects without requiring Excel
    - Integration test approach: live Excel sessions marked `@pytest.mark.excel`
    - CI configuration: how `@pytest.mark.excel` tests are excluded from automated pipelines
    - Test naming conventions: `test_should_<expected_behavior>_when_<condition>`
    - Fixtures and helpers provided by the `ezxl` test package
    - Running the test suite locally with and without a live Excel instance

    In the meantime, run the existing tests with:

    ```bash
    # All tests (unit only — no Excel required)
    pytest

    # With coverage
    pytest --cov=ezxl --cov-report=term-missing

    # Include integration tests (requires Excel installed and running)
    pytest -m excel
    ```
