from __future__ import annotations

from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_initial_application_state_renders_without_errors() -> None:
    app = AppTest.from_file("app.py", default_timeout=15)

    app.run()

    assert not app.exception
    rendered_markdown = "\n".join(str(element.value) for element in app.get("markdown"))
    assert "SEC Non-GAAP" in rendered_markdown
    assert "Peer Benchmarking" in rendered_markdown
    assert "Built by Vibhor" in rendered_markdown
    assert any(
        "Start with an issuer search" in element.value
        for element in app.get("subheader")
    )


def test_apple_theme_declares_forced_light_contrast_for_core_surfaces() -> None:
    source = Path("app.py").read_text()

    assert "APPLE_THEME_CSS" in source
    assert "--apple-blue: #0071E3" in source
    assert "color-scheme: light !important" in source
    assert '[data-testid="stSidebar"]' in source
    assert '[data-testid="stAlert"]' in source
    assert '[data-testid="stDataFrame"]' in source
    assert '[data-testid="stNumberInput"]' in source
    assert "--apple-sidebar: #1D1D1F" in source
