import requests

url = "https://script.googleusercontent.com/macros/echo?user_content_key=Hp0x30o9nYoyE-zsQF8fUXjmNFjd94xslVdB82ppPAnL8ndYjcnIOJobETxlv8pRAaj1XE-jJZ6T7sZXuY7HxpSSj25Y-6jam5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnB_bXk1Fp1xyXUCi7YJIQlwHpwAzj2FSCLl3gk5mqg704GAXFQ7NFlaHVXqAY8piubhgzlB--lg4mhqBhoXIImGBs-fWFaZ6uw&lib=M9zgSB9AZuUM-ZH-LYbUIysmBSqTSfeVE"

response = requests.get(url)
if response.status_code == 200:
    data = response.json()
    print(data)
else:
    print("Error... System")