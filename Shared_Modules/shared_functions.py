# =====================For All=====================
def line_notify(message = '', token = 'u7bfH6ad2gDcHvvPrtHR9sjJ8AYmQ7tNl0VBf7piO4q', notify_ornot = True):
    import requests
    if not notify_ornot:
        return
    # HTTP header 參數設定
    headers = { "Authorization": "Bearer " + token }
    data = { 'message': message }
    # 以 requests 發送 POST 請求
    requests.post("https://notify-api.line.me/api/notify",
    headers = headers, data = data)