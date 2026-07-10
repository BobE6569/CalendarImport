from pathlib import Path

from calendar_import.timezone_resolver import TimeZoneResolver, mapping_path_for_workbook


def test_prompted_time_zone_is_saved_to_text_file(tmp_path):
    mapping_path = tmp_path / "Trip_timezones.txt"
    resolver = TimeZoneResolver(mapping_path, lambda location, fallback: "Europe/Oslo")

    assert resolver.resolve("Bergen") == "Europe/Oslo"

    saved = mapping_path.read_text(encoding="utf-8")
    assert "Bergen=Europe/Oslo" in saved

    next_resolver = TimeZoneResolver(mapping_path)
    assert next_resolver.resolve("Bergen") == "Europe/Oslo"


def test_mapping_path_for_workbook_uses_workbook_name(tmp_path):
    workbook_path = Path("Norway Itinerary Aug 2026 Sub.xlsx")

    assert mapping_path_for_workbook(tmp_path, workbook_path) == tmp_path / "Norway_Itinerary_Aug_2026_Sub_timezones.txt"
