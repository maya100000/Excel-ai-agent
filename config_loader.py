import json
import os

def load_config(config_path: str = None) -> dict:
    """
    Load a client config JSON file.
    If no path given, loads client_config.json from the same folder.
    """
    if config_path is None:
        config_path = os.path.join(os.path.dirname(__file__), "client_config.json")

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

def get_weekend_days(config: dict) -> set:
    return set(config["work_week"]["weekend"])

def get_workdays(config: dict) -> set:
    return set(config["work_week"]["workdays"])

def get_holiday_events(config: dict) -> set:
    return set(config["business_events"]["holiday"])

def get_high_volume_events(config: dict) -> set:
    return set(config["business_events"]["high_volume"])

def get_sheet_translation(config: dict) -> dict:
    return config["sheet_translation"]

def get_column_mapping(config: dict) -> dict:
    return config["column_mapping"]

def get_day_translation(config: dict) -> dict:
    return config["day_translation"]

def get_notes_translation(config: dict) -> dict:
    return config["notes_translation"]

def get_month_order(config: dict) -> list:
    return config["month_order"]

def get_highlight_colors(config: dict) -> dict:
    return config["output"]["highlight_colors"]

def get_llm_settings(config: dict) -> dict:
    return config["llm"]

def get_anomaly_settings(config: dict) -> dict:
    return config["anomaly_detection"]

def get_agent_sheet_settings(config: dict) -> dict:
    return config["agent_sheet"]

def describe_config(config: dict) -> str:
    """Return a human-readable summary of the config for logging."""
    lines = [
        f"Client: {config['client_name']} (ID: {config['client_id']})",
        f"Domain: {config['domain']}",
        f"Language: {config['language']}",
        f"Work week: {', '.join(config['work_week']['workdays'])}",
        f"Weekend: {', '.join(config['work_week']['weekend'])}",
        f"Sheets mapped: {len(config['sheet_translation'])}",
        f"Columns mapped: {len(config['column_mapping'])}",
        f"KPIs defined: {len(config['kpis'])}",
        f"Business events: {len(config['notes_translation'])}",
    ]
    return "\n".join(lines)

if __name__ == "__main__":
    cfg = load_config()
    print(describe_config(cfg))
