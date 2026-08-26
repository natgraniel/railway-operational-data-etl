from src.extractors.pdf_extractor import (
    _find_shared_registrations,
    _parse_useful_commercial_line,
)

def test_finds_registration_shared_by_adjacent_commercial_lines():
    token_lines = [
        [
            "05h45",
            "/06h47",
            "06h45",
            "/07h47",
            "401-402",
            "Cancún-P",
            "Carmen-Cancún-P",
        ],
        ["Cancún", "TS09", "D009", "-"],
        [
            "09h39/11h00",
            "10h39/12h00",
            "403-404",
            "Carmen-Cancún",
        ],
    ]

    result = _find_shared_registrations(token_lines)

    assert result == {
        0: "D009",
        2: "D009",
    }

def test_parses_useful_commercial_train_as_reserve():
    tokens = ["Cancún", "TS03", "D003", "-"]

    result = _parse_useful_commercial_line(
        tokens,
        " ".join(tokens),
    )

    assert result is not None
    assert result.workshop_station == "Cancún"
    assert result.registration == "D003"
    assert result.status == "RESERVA"


def test_ignores_maintenance_row_in_useful_commercial_section():
    tokens = [
        "Cancún",
        "TS08",
        "D008",
        "28-ago-26",
        "Mantenimiento",
        "correctivo",
    ]

    result = _parse_useful_commercial_line(
        tokens,
        " ".join(tokens),
    )

    assert result is None    