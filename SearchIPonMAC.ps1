#Search IP on MAC
arp -a | select-string "3c-ef-8c" |% { $_.ToString().Trim().Split(" ")[0] }