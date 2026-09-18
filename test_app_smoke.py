from __future__ import annotations

from streamlit.testing.v1 import AppTest


def test_initial_application_state_renders_without_errors() -> None:
    app = AppTest.from_file("app.py", default_timeout=15)

    app.run()

    assert not app.exception
    assert any(
        "Start with an issuer search" in element.value
        for element in app.get("subheader")
    )
