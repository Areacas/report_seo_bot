import os


def creat_directories(user_id):
    directories = ["reports_sku", "reports_group", "ready_reports_sku", "ready_reports_group"]

    for directory in directories:
        if not os.path.exists(f"reports/{user_id}/{directory}"):
            os.makedirs(f"reports/{user_id}/{directory}")
