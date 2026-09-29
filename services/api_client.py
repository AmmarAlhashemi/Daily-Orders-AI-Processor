import requests


class APIClient:
    def __init__(self, base_url="https://daily-orders-ai-backend.onrender.com"):
        self.base_url = base_url.rstrip("/")

    def parse_orders(self, raw_data: str, use_fallback: bool = False):
        response = requests.post(
            f"{self.base_url}/parse-orders",
            json={
                "raw_data": raw_data,
                "use_fallback": use_fallback,
            },
            timeout=120,
        )

        if response.status_code == 503:
            error_data = response.json()
            raise Exception(error_data["detail"]["code"])

        response.raise_for_status()

        return response.json()

    def smart_split(self, excel_data: list, use_fallback: bool = False):
        response = requests.post(
            f"{self.base_url}/smart-split",
            json={
                "excel_data": excel_data,
                "use_fallback": use_fallback,
            },
            timeout=120,
        )

        if response.status_code == 503:
            error_data = response.json()
            raise Exception(error_data["detail"]["code"])

        response.raise_for_status()

        return response.json()
