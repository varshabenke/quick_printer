import qrcode

website_url = "https://overtime-unwanted-deflation.ngrok-free.dev"

img = qrcode.make(website_url)

img.save("start_qr.png")

print("QR Generated Successfully!")







