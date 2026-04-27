from pathlib import Path

import server
from pptx import Presentation


COMMAND_MAP = Path("command_map.json")


def _payload(response):
    data = response.get_json(silent=True)
    assert isinstance(data, dict), response.data
    assert "success" in data, data
    assert "status" in data, data
    assert "intent" in data, data
    assert "message" in data, data
    return data


def _assert_office_file(data, suffix):
    assert data["success"] is True, data
    assert data["intent"] == "office_automation", data
    assert data.get("file_path"), data
    path = Path(data["file_path"])
    assert path.exists(), data
    assert path.suffix.lower() == suffix, data
    assert path.stat().st_size > 0, data


def main():
    original_command_map = COMMAND_MAP.read_bytes() if COMMAND_MAP.exists() else None

    # Do not launch real desktop apps or open blocking file pickers in smoke tests.
    server.system_core.open_path = lambda *args, **kwargs: True
    server.ui.manual_selector = lambda: ""

    try:
        client = server.app.test_client()

        cases = [
            ("create a new Excel file", ".xlsx"),
            ("create a new Word document", ".docx"),
            ("create a new PowerPoint presentation", ".pptx"),
            ("make a spreadsheet with 3 columns and 5 rows", ".xlsx"),
            ("create a presentation with 3 slides about sales performance", ".pptx"),
        ]

        generated = {}
        for command, suffix in cases:
            data = _payload(client.post("/execute", json={"command": command}))
            _assert_office_file(data, suffix)
            assert not data.get("requires_manual_selection"), data
            generated[command] = Path(data["file_path"])

        three_slide_path = generated["create a presentation with 3 slides about sales performance"]
        assert len(Presentation(str(three_slide_path)).slides) == 3

        deck = _payload(client.post(
            "/office/execute",
            json={"app": "powerpoint", "raw": "create a presentation with 2 slides"},
        ))
        _assert_office_file(deck, ".pptx")
        deck_path = deck["file_path"]

        first = _payload(client.post(
            "/office/execute",
            json={
                "app": "powerpoint",
                "raw": "write First in title on slide 1",
                "file_path": deck_path,
            },
        ))
        assert first["success"] is True, first

        second = _payload(client.post(
            "/office/execute",
            json={
                "app": "powerpoint",
                "raw": "write Second in title on slide 2",
                "file_path": deck_path,
            },
        ))
        assert second["success"] is True, second

        prs = Presentation(deck_path)
        first_text = " ".join(shape.text for shape in prs.slides[0].shapes if hasattr(shape, "text"))
        second_text = " ".join(shape.text for shape in prs.slides[1].shapes if hasattr(shape, "text"))
        assert "first" in first_text.lower(), first_text
        assert "second" in second_text.lower(), second_text

        unknown = _payload(client.post("/execute", json={"command": "open someunknownapp"}))
        assert unknown["success"] is False, unknown
        assert unknown["intent"] == "app_launch", unknown
        assert unknown.get("requires_manual_selection") is True, unknown

        invalid = _payload(client.post("/execute", json={"command": "dance please"}))
        assert invalid["success"] is False, invalid
        assert invalid["error_code"] == "UNKNOWN_COMMAND", invalid

        page = client.get("/")
        assert page.status_code == 200
        assert b"/static/reliability.js" in page.data

        frontend = client.get("/static/reliability.js")
        assert frontend.status_code == 200
        assert b"finally" in frontend.data
        assert b"Unexpected backend response" in frontend.data

    finally:
        if original_command_map is None:
            if COMMAND_MAP.exists():
                COMMAND_MAP.unlink()
        else:
            COMMAND_MAP.write_bytes(original_command_map)

    print("Office route smoke tests passed.")


if __name__ == "__main__":
    main()
