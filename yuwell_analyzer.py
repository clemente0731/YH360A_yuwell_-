
import struct
import csv
import math
import statistics
import os
import argparse
from datetime import datetime

# ==========================================
# Part 1: BYS File Parser
# ==========================================

def parse_bys_file(filename, output_csv=None):
    """
    Parses a Yuwell .BYS file and extracts session records.
    Returns a list of dictionaries containing the parsed data.
    """
    print(f"Parsing {filename}...")
    
    if not os.path.exists(filename):
        print(f"Error: File {filename} not found.")
        return []

    records = []
    
    try:
        with open(filename, 'rb') as f:
            content = f.read()
            
        # Search for records based on the magic marker
        # Pattern: 28 46 96 28
        magic = b'\x28\x46\x96\x28'
        
        offset = 0
        while True:
            idx = content.find(magic, offset)
            if idx == -1:
                break
                
            # Check if we have enough bytes for a full record
            # Record size = 30 bytes
            # [Marker: 4] [Data: 12] [TS1: 6] [TS2: 6] [Spacer: 2]
            
            if idx + 30 > len(content):
                break
                
            record_bytes = content[idx : idx+30]
            
            # Parse fields
            # Data: 12 bytes (6 x uint16 BE)
            data_raw = record_bytes[4:16]
            data_values = struct.unpack('>6H', data_raw)
            
            # TS1: 6 bytes
            ts1_raw = record_bytes[16:22]
            try:
                ts1_str = f"20{ts1_raw[0]:02d}-{ts1_raw[1]:02d}-{ts1_raw[2]:02d} {ts1_raw[3]:02d}:{ts1_raw[4]:02d}:{ts1_raw[5]:02d}"
                # Verify date validity
                datetime.strptime(ts1_str, '%Y-%m-%d %H:%M:%S')
            except:
                ts1_str = "Invalid Date"
            
            # TS2: 6 bytes
            ts2_raw = record_bytes[22:28]
            try:
                ts2_str = f"20{ts2_raw[0]:02d}-{ts2_raw[1]:02d}-{ts2_raw[2]:02d} {ts2_raw[3]:02d}:{ts2_raw[4]:02d}:{ts2_raw[5]:02d}"
                datetime.strptime(ts2_str, '%Y-%m-%d %H:%M:%S')
            except:
                ts2_str = "Invalid Date"
            
            # Spacer
            spacer = struct.unpack('>H', record_bytes[28:30])[0]
            
            # Duration (min) is at index 5
            duration = data_values[5]
            
            # Pressure Raw is at index 3
            pressure_raw = data_values[3]
            
            record = {
                'Offset': f"0x{idx:X}",
                'Start Time': ts1_str,
                'End Time': ts2_str,
                'Duration (min)': duration,
                'Pressure Raw': pressure_raw,
                'Val1': data_values[0],
                'Val2': data_values[1],
                'Val3': data_values[2],
                'Val4': data_values[4],
                'Spacer': f"0x{spacer:04X}"
            }
            records.append(record)
            
            offset = idx + 4 # Move past this marker
            
        print(f"Found {len(records)} records.")
        
        # Write to CSV if requested
        if output_csv and records:
            keys = records[0].keys()
            with open(output_csv, 'w', newline='', encoding='utf-8') as f_csv:
                dict_writer = csv.DictWriter(f_csv, keys)
                dict_writer.writeheader()
                dict_writer.writerows(records)
            print(f"Written to {output_csv}")
            
        return records

    except Exception as e:
        print(f"Error parsing file: {e}")
        return []

# ==========================================
# Part 2: Data Analysis & Visualization
# ==========================================

def calculate_stats(values, name):
    if not values:
        print(f"--- {name} ---\nNo data available.\n")
        return
    
    _min = min(values)
    _max = max(values)
    _avg = sum(values) / len(values)
    _stdev = statistics.stdev(values) if len(values) > 1 else 0
    
    print(f"--- {name} ---")
    print(f"Range: {_min:.2f} - {_max:.2f}")
    print(f"Mean:  {_avg:.2f}")
    print(f"StdDev:{_stdev:.2f}")
    # Print first 10 samples
    samples = [f"{v:.2f}" for v in values[:10]]
    print(f"Samples: {', '.join(samples)} ...")
    print("")

def generate_html_report(data, filename="report.html"):
    print(f"\nGenerating HTML report: {filename}...")
    
    # Sort data by start time to ensure charts are chronological
    data.sort(key=lambda x: x['Start Time'])
    
    dates = [d['Start Time'].strftime('%Y-%m-%d') for d in data]
    durations = [d['Duration'] for d in data]
    pressures = [d['Pressure'] for d in data]
    
    # Calculate AHI (Val3 / Hour)
    ahi = []
    for d in data:
        hours = d['Duration'] / 60.0
        if hours > 0.5:
            ahi.append(d['Val3'] / hours)
        else:
            ahi.append(0)
            
    # Calculate Leak (Val2 / 100)
    leak = []
    for d in data:
        leak.append(d['Val2'] / 100.0)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Yuwell CPAP Data Analysis</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body {{ font-family: "Microsoft YaHei", sans-serif; margin: 20px; color: #333; }}
            .chart-container {{ width: 80%; margin: 20px auto; background: #fff; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }}
            .explanation {{ background: #f8f9fa; padding: 20px; border-left: 5px solid #007bff; margin: 20px auto; width: 80%; border-radius: 4px; }}
            h1 {{ text-align: center; color: #2c3e50; }}
            h2 {{ border-bottom: 2px solid #eaeaea; padding-bottom: 10px; margin-top: 0; color: #007bff; }}
            ul {{ line-height: 1.6; }}
            .stats-table {{ width: 80%; margin: 20px auto; border-collapse: collapse; }}
            .stats-table th, .stats-table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            .stats-table th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h1>鱼跃呼吸机数据分析报告 (Yuwell CPAP Analysis)</h1>
        <p style="text-align:center;">生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        
        <div class="explanation">
            <h2>1. 使用时长 (Duration)</h2>
            <p>此图表显示了每次治疗的持续时间（分钟）。</p>
            <ul>
                <li><strong>正常标准:</strong> 一般建议每晚使用超过 4 小时 (240分钟)。</li>
                <li><strong>趋势分析:</strong> 观察是否有中断治疗或使用时间逐渐增加的趋势。</li>
            </ul>
        </div>
        <div class="chart-container">
            <canvas id="durationChart"></canvas>
        </div>

        <div class="explanation">
            <h2>2. 治疗压力 (Treatment Pressure)</h2>
            <p>此图表显示了每次治疗的平均压力值 (cmH2O)。</p>
            <ul>
                <li><strong>临床意义:</strong> 反映了气道阻塞的严重程度。压力越高，说明需要更强的气流来保持气道通畅。</li>
                <li><strong>稳定性:</strong> 稳定的压力通常意味着病情控制良好。</li>
            </ul>
        </div>
        <div class="chart-container">
            <canvas id="pressureChart"></canvas>
        </div>

        <div class="explanation">
            <h2>3. 呼吸暂停指数 (AHI - Estimated)</h2>
            <p>此图表显示了推算的每小时呼吸暂停/低通气次数 (Val3 / 小时)。</p>
            <ul>
                <li><strong>正常值:</strong> &lt; 5 次/小时</li>
                <li><strong>轻度:</strong> 5-15 次/小时</li>
                <li><strong>中度:</strong> 15-30 次/小时</li>
                <li><strong>重度:</strong> &gt; 30 次/小时</li>
                <li><strong>注意:</strong> 这是基于设备记录推算的估算值，可能与PSG多导睡眠监测结果有出入。</li>
            </ul>
        </div>
        <div class="chart-container">
            <canvas id="ahiChart"></canvas>
        </div>
        
        <div class="explanation">
            <h2>4. 漏气量 (Leak Rate - Estimated)</h2>
            <p>此图表显示了推算的平均漏气量 (Val2 / 100 L/min)。</p>
            <ul>
                <li><strong>正常范围:</strong> 0 - 60 L/min</li>
                <li><strong>注意:</strong> 过高的漏气量可能会影响治疗效果和数据准确性。如持续过高，请检查面罩佩戴是否紧密。</li>
            </ul>
        </div>
        <div class="chart-container">
            <canvas id="leakChart"></canvas>
        </div>

        <script>
            const dates = {dates};
            
            new Chart(document.getElementById('durationChart'), {{
                type: 'line',
                data: {{
                    labels: dates,
                    datasets: [{{
                        label: 'Duration (min)',
                        data: {durations},
                        borderColor: 'rgb(75, 192, 192)',
                        backgroundColor: 'rgba(75, 192, 192, 0.2)',
                        tension: 0.1,
                        fill: true
                    }}]
                }},
                options: {{ scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Minutes' }} }} }} }}
            }});
            
            new Chart(document.getElementById('pressureChart'), {{
                type: 'line',
                data: {{
                    labels: dates,
                    datasets: [{{
                        label: 'Pressure (cmH2O)',
                        data: {pressures},
                        borderColor: 'rgb(255, 99, 132)',
                        backgroundColor: 'rgba(255, 99, 132, 0.2)',
                        tension: 0.1,
                        fill: false
                    }}]
                }},
                options: {{ scales: {{ y: {{ beginAtZero: false, title: {{ display: true, text: 'cmH2O' }} }} }} }}
            }});
            
            new Chart(document.getElementById('ahiChart'), {{
                type: 'bar',
                data: {{
                    labels: dates,
                    datasets: [{{
                        label: 'AHI (Events/Hr)',
                        data: {ahi},
                        backgroundColor: 'rgb(54, 162, 235)',
                    }}]
                }},
                options: {{ scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'Events / Hour' }} }} }} }}
            }});
            
            new Chart(document.getElementById('leakChart'), {{
                type: 'line',
                data: {{
                    labels: dates,
                    datasets: [{{
                        label: 'Leak (L/min)',
                        data: {leak},
                        borderColor: 'rgb(153, 102, 255)',
                        backgroundColor: 'rgba(153, 102, 255, 0.2)',
                        tension: 0.1,
                        fill: true
                    }}]
                }},
                options: {{ scales: {{ y: {{ beginAtZero: true, title: {{ display: true, text: 'L/min' }} }} }} }}
            }});
        </script>
    </body>
    </html>
    """
    
    with open(filename, "w", encoding='utf-8') as f:
        f.write(html_content)
    print(f"Report generated: {filename}")

def run_analysis(records):
    """
    Analyzes the extracted records and prints statistics.
    """
    # Convert records to analysis format
    data = []
    for r in records:
        try:
            if r['Start Time'] == "Invalid Date": continue
            
            item = {
                'Start Time': datetime.strptime(r['Start Time'], '%Y-%m-%d %H:%M:%S'),
                'Duration': float(r['Duration (min)']),
                'Pressure': float(r['Pressure Raw']) / 1000.0,
                'Val1': float(r['Val1']),
                'Val2': float(r['Val2']),
                'Val3': float(r['Val3']),
                'Val4': float(r['Val4']),
            }
            data.append(item)
        except ValueError:
            continue
            
    print(f"\nAnalyzing {len(data)} valid records...\n")
    
    # Basic Stats
    calculate_stats([d['Duration'] for d in data], "Duration (min)")
    calculate_stats([d['Pressure'] for d in data], "Pressure (cmH2O)")
    
    # Hypothesis Testing Stats
    print("--- Derived Metrics ---")
    
    # AHI
    ahi_vals = []
    for d in data:
        hours = d['Duration'] / 60.0
        if hours > 0.5:
            ahi_vals.append(d['Val3'] / hours)
    calculate_stats(ahi_vals, "AHI (Val3 / Hour)")
    
    # Leak
    leak_vals = [d['Val2']/100.0 for d in data if d['Duration'] > 30]
    calculate_stats(leak_vals, "Leak Rate (Val2 / 100)")
    
    return data

# ==========================================
# Part 3: Main Entry Point
# ==========================================

def main():
    parser = argparse.ArgumentParser(description="Yuwell CPAP Data Analyzer")
    parser.add_argument("input_file", nargs='?', default="YHSD-NEW.BYS", help="Path to .BYS file")
    parser.add_argument("--csv", default="summary.csv", help="Output CSV filename")
    parser.add_argument("--html", default="report.html", help="Output HTML report filename")
    
    args = parser.parse_args()
    
    # 1. Parse
    records = parse_bys_file(args.input_file, args.csv)
    
    if not records:
        print("No records found. Exiting.")
        return

    # 2. Analyze
    data = run_analysis(records)
    
    # 3. Generate Report
    if data:
        generate_html_report(data, args.html)

if __name__ == "__main__":
    main()
