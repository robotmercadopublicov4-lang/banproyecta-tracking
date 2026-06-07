import smtplib

gmail = "robotmercadopublicov4@gmail.com"
password = "fkejbgvxrgnunbbq"

try:

    server = smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465
    )

    server.login(gmail, password)

    print("LOGIN OK")

    server.quit()

except Exception as e:

    print(e)