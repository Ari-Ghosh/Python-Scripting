import csv
import subprocess
import time

interval = 1

output_file = "GPU_MONITORED_DATA.csv"
with open(output_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Timestamp", "GPU %", "GPU Memory Used"])

while True:
    timestamp = time.time()
    try:
        output = subprocess.check_output(["nvidia-smi", "--query-gpu=utilization.gpu,memory.used", "--format=csv,noheader,nounits"])
        gpu_percent, gpu_memory = output.decode().strip().split(",")
    except subprocess.CalledProcessError:
        gpu_percent, gpu_memory = None, None

    if gpu_percent is not None:
        print(f"GPU: {gpu_percent}% ({gpu_memory} memory used)")
    else:
        print("GPU: not available")

    with open(output_file, "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([timestamp, gpu_percent, gpu_memory])

    time.sleep(interval)
