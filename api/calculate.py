"""
Vercel Python Serverless Function
POST /api/calculate
  - multipart/form-data: file (xlsx), start (YYYY-MM-DD), end (YYYY-MM-DD)
  - returns JSON: { summary, total_receivable, total_income, contract_count, files:{lease,single,income} }
"""
from http.server import BaseHTTPRequestHandler
import cgi
import json
import os
import sys
import uuid
import base64
import shutil
import traceback
import tempfile

# Add lib/ to path so we can import lease_calculator
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'lib'))

from lease_calculator import LeaseCalculator


def _read_b64(path: str) -> str:
    """Read a file and return base64-encoded string."""
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')


def _find_output_files(out_dir: str):
    """Find the generated xlsx files in out_dir (timestamped names)."""
    files = {}
    for fname in os.listdir(out_dir):
        if fname.endswith('-lease.xlsx') or fname == 'lease.xlsx':
            files['lease'] = os.path.join(out_dir, fname)
        elif fname.endswith('-single.xlsx') or fname == 'single.xlsx':
            files['single'] = os.path.join(out_dir, fname)
        elif fname.endswith('-income.xlsx') or fname == 'income.xlsx':
            files['income'] = os.path.join(out_dir, fname)
    # Also check for files matching pattern YYYYMMDDHHMMSS-xxx.xlsx
    for fname in os.listdir(out_dir):
        if fname.endswith('.xlsx'):
            lower = fname.lower()
            if 'lease' in lower and 'lease' not in files:
                files['lease'] = os.path.join(out_dir, fname)
            elif 'single' in lower and 'single' not in files:
                files['single'] = os.path.join(out_dir, fname)
            elif 'income' in lower and 'income' not in files:
                files['income'] = os.path.join(out_dir, fname)
    return files


class handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        """Suppress default request logging."""
        pass

    def _send_json(self, status: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        tmp_dir = None
        try:
            # Parse multipart/form-data
            content_type = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in content_type:
                self._send_json(400, {'error': '请求必须是 multipart/form-data 格式'})
                return

            environ = {
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE': content_type,
                'CONTENT_LENGTH': self.headers.get('Content-Length', '0'),
            }
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ=environ,
            )

            # Extract file
            if 'file' not in form:
                self._send_json(400, {'error': '缺少文件字段 "file"'})
                return

            file_item = form['file']
            file_data = file_item.file.read()
            if not file_data:
                self._send_json(400, {'error': '上传的文件为空'})
                return

            # Extract dates
            start_val = form.getvalue('start', '')
            end_val = form.getvalue('end', '')
            if not start_val or not end_val:
                self._send_json(400, {'error': '缺少 start 或 end 参数（格式 YYYY-MM-DD）'})
                return

            # Write input file to /tmp
            tmp_dir = tempfile.mkdtemp(prefix='lease_', dir='/tmp')
            input_path = os.path.join(tmp_dir, 'input.xlsx')
            out_dir = os.path.join(tmp_dir, 'output')
            os.makedirs(out_dir, exist_ok=True)

            with open(input_path, 'wb') as f:
                f.write(file_data)

            # Run calculation
            calc = LeaseCalculator(input_path)
            result_dfs = calc.process_all_contracts(
                start_month=start_val,
                end_month=end_val,
                output_dir=out_dir,
                enable_log=False,
                aux_columns=False,
            )

            # Find output files
            output_files = _find_output_files(out_dir)

            missing = [k for k in ('lease', 'single', 'income') if k not in output_files]
            if missing:
                self._send_json(500, {
                    'error': f'计算完成但未找到输出文件: {missing}。out_dir 内容: {os.listdir(out_dir)}'
                })
                return

            # Build summary from results
            # process_all_contracts returns (summary_df, monthly_receivables_df, monthly_income_df)
            summary = []
            total_receivable = 0.0
            total_income = 0.0

            summary_df = result_dfs[0] if result_dfs else None
            if summary_df is not None and len(summary_df) > 0:
                for _, row in summary_df.iterrows():
                    recv = float(row.get('应收总额', 0) or 0)
                    inc = float(row.get('收入总额', 0) or 0)
                    bank = float(row.get('银行对账单', 0) or 0)
                    inv = float(row.get('发票对账', 0) or 0)
                    total_receivable += recv
                    total_income += inc
                    summary.append({
                        'customer': str(row.get('客户名称', '')),
                        'merchant_id': str(row.get('商户编号', '')),
                        'receivable': round(recv, 2),
                        'income': round(inc, 2),
                        'bank_matched': round(bank, 2),
                        'invoice_matched': round(inv, 2),
                        'notes': str(row.get('数据备注', '') or ''),
                    })

            response = {
                'contract_count': len(summary),
                'total_receivable': round(total_receivable, 2),
                'total_income': round(total_income, 2),
                'summary': summary,
                'files': {
                    'lease': _read_b64(output_files['lease']),
                    'single': _read_b64(output_files['single']),
                    'income': _read_b64(output_files['income']),
                },
            }

            self._send_json(200, response)

        except Exception as e:
            tb = traceback.format_exc()
            self._send_json(500, {
                'error': f'计算失败: {str(e)}',
                'traceback': tb,
            })
        finally:
            # Clean up temp files
            if tmp_dir and os.path.exists(tmp_dir):
                try:
                    shutil.rmtree(tmp_dir)
                except Exception:
                    pass
