#Search IP on MAC
arp -a | select-string "90-09-d0-60" |% { $_.ToString().Trim().Split(" ")[0] }