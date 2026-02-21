"""
generate_sample_data.py
Generate sample DHT11 temperature/humidity logger data as an Excel file.
"""
import openpyxl
from datetime import datetime, timedelta
import random

def generate_sample_data(filename="dht11_log_sample.xlsx", rows=200):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Raw Data"

    # Headers matching typical N5110 DHT11 logger output
    headers = ["#", "Date", "Time", "Temperature (°C)", "Humidity (%)", "Heat Index (°C)"]
    ws.append(headers)

    start_dt = datetime(2024, 1, 1, 0, 0, 0)
    temp_base = 22.0
    hum_base = 55.0

    for i in range(1, rows + 1):
        dt = start_dt + timedelta(minutes=15 * i)
        # Simulate realistic temperature and humidity with daily variation
        hour = dt.hour
        temp = round(temp_base + 5 * (hour / 12 if hour <= 12 else (24 - hour) / 12) + random.uniform(-1.5, 1.5), 1)
        hum  = round(hum_base - 10 * (hour / 12 if hour <= 12 else (24 - hour) / 12) + random.uniform(-3, 3), 1)
        temp = max(10.0, min(45.0, temp))
        hum  = max(20.0, min(95.0, hum))
        # Simple heat index approximation
        hi   = round(temp + 0.1 * hum, 1)

        ws.append([i, dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M:%S"), temp, hum, hi])

    wb.save(filename)
    print(f"Sample data saved to: {filename}")
    return filename


if __name__ == "__main__":
    generate_sample_data()
