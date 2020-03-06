import requests

userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"
header = {
    # "origin": "https://www.tianyancha.com",
    "Referer": "Origin: https://www.tianyancha.com",
    'User-Agent': userAgent,
}



def main_spider(account, password):
    postUrl = "https://www.tianyancha.com/cd/login.json"
    postData = {
        "passport": account,
        "password": password,
    }
    responseRes = requests.post(postUrl, data=postData, headers=header)
    # 无论是否登录成功，状态码一般都是 statusCode = 200
    print(f"statusCode = {responseRes.status_code}")
    print(f"text = {responseRes.text}")


if __name__ == "__main__":
    # 从返回结果来看，有登录成功
    main_spider("15977469433", "xyl666xyl666")