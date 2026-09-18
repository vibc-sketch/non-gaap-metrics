from __future__ import annotations

from streamlit.testing.v1 import AppTest


def test_initial_application_state_renders_without_errors() -> None:
    app = AppTest.from_file("app.py", default_timeout=15)

    app.run()

    assert not app.exception
    rendered_markdown = "\n".join(str(element.value) for element in app.get("markdown"))
    assert "SEC Non-GAAP Intelligence" in rendered_markdown
    assert "Built by Vibhor" in rendered_markdown
    assert any(
        "Start with an issuer search" in element.value
        for element in app.get("subheader")
    )
