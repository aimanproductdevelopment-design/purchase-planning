"""
APPS — Full Pipeline Runner
รัน: Export JST → คำนวณ → ออก Excel ในขั้นตอนเดียว

วิธีใช้:
  python run_full_pipeline.py             # รันปกติ (browser ซ่อน)
  python run_full_pipeline.py --show      # เปิด browser ให้เห็น
  python run_full_pipeline.py --skip-export  # ข้ามส่วน export (ใช้ data เดิม)

ตั้งเวลาอัตโนมัติทุกเช้า 08:00 (cron):
  0 8 * * * cd /Users/viit/apps && python3 auto_export/run_full_pipeline.py >> logs/daily.log 2>&1
"""
import sys
import os
import argparse
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--show",         action="store_true", help="แสดง browser ขณะ export")
    parser.add_argument("--skip-export",  action="store_true", help="ข้าม JST export ใช้ data เดิม")
    parser.add_argument("--output",       default="output/purchasing_plan.xlsx", help="ไฟล์ output")
    args = parser.parse_args()

    log_dir = Path(__file__).parent.parent / "logs"
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f"pipeline_{datetime.now().strftime('%Y%m%d')}.log"

    print(f"🚀 APPS Full Pipeline — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"   Log: {log_file}")

    if not args.skip_export:
        print("\n── Step 1: Export จาก JST ERP ──────────────────")
        from auto_export.jst_exporter import run as export_run
        export_run(headless=not args.show)
    else:
        print("\n── Step 1: ข้าม Export (ใช้ data เดิม) ──────────")

    print("\n── Step 2: คำนวณ Forecast / Safety Stock / Alert ──")
    os.chdir(Path(__file__).parent.parent)
    from run_planning import run as plan_run
    results = plan_run(output_path=args.output)

    print(f"\n✅ เสร็จสิ้น — ผลลัพธ์: {args.output}")
    return results


if __name__ == "__main__":
    main()
