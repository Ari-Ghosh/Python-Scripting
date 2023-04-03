import csv
import psutil
import time

interval = 5

output_file = "CPU_MONITORED_DATA.csv"

with open(output_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Timestamp", "CPU %", "Cores Used"])

while True:
    timestamp = time.time()

    cpu_percent = psutil.cpu_percent()
    cores_used = psutil.cpu_count(logical=False)

    print(f"CPU: {cpu_percent:.1f}% ({cores_used} cores used)")

    with open(output_file, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([timestamp, cpu_percent, cores_used])

    time.sleep(interval)
