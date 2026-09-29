import os
from datetime import datetime, timedelta
from ui.app import App
from config import MASTER_AC_DIR
from services.rollover_service import perform_monthly_rollover_and_archive

def check_and_run_monthly_rollover():
    """Checks whether previous month archiving/rollover is needed on startup."""
    today = datetime.today().date()
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - timedelta(days=1)

    prev_month_folder_name = last_day_prev_month.strftime("%b-%Y").upper()

    prev_month_archive_folder = MASTER_AC_DIR / prev_month_folder_name

    if MASTER_AC_DIR.exists() and not prev_month_archive_folder.exists():
        perform_monthly_rollover_and_archive(today)

if __name__ == "__main__":
    check_and_run_monthly_rollover()
    app = App()
    app.mainloop()