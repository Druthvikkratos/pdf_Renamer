"""
Background job runner for processing large PDF batches without holding
everything in memory. Each job streams uploaded PDFs to disk, processes
them one at a time, and writes results straight into on-disk zip files.
"""
import os
import tempfile
import threading
import uuid
import zipfile

from .renamer import (
    load_excel, ExcelParseError, read_pdf_text, extract_doc_no_from_filename,
    extract_all_numbers, find_doc_no_matches_in_content, build_new_filename,
)

JOBS = {}
JOBS_LOCK = threading.Lock()
JOBS_DIR = os.path.join(tempfile.gettempdir(), "pdfrenamer_jobs")
os.makedirs(JOBS_DIR, exist_ok=True)


class Job:
    def __init__(self, job_id: str, total: int):
        self.id = job_id
        self.total = total
        self.processed = 0
        self.current_filename = ""
        self.success = []
        self.failed = []
        self.excel_error = None
        self.excel_warnings = []
        self.excel_total_rows = 0
        self.excel_matched_count = 0
        self.excel_unmatched = []
        self.done = False
        self.error = None

        self.dir = os.path.join(JOBS_DIR, job_id)
        os.makedirs(self.dir, exist_ok=True)
        self.upload_dir = os.path.join(self.dir, "uploads")
        os.makedirs(self.upload_dir, exist_ok=True)
        self.success_zip_path = os.path.join(self.dir, "renamed_success.zip")
        self.failed_zip_path = os.path.join(self.dir, "failed_pdfs.zip")

    def status_dict(self):
        return {
            "job_id": self.id,
            "total": self.total,
            "processed": self.processed,
            "current_filename": self.current_filename,
            "done": self.done,
            "error": self.error,
        }

    def result_dict(self):
        return {
            "excel_error": self.excel_error,
            "excel_warnings": self.excel_warnings,
            "excel_total_rows": self.excel_total_rows,
            "excel_matched_count": self.excel_matched_count,
            "excel_unmatched": self.excel_unmatched,
            "success": self.success,
            "failed": self.failed,
            "total_pdfs": self.total,
            "has_success_zip": os.path.exists(self.success_zip_path),
            "has_failed_zip": os.path.exists(self.failed_zip_path),
        }


def create_job(total: int) -> Job:
    job_id = str(uuid.uuid4())
    job = Job(job_id, total)
    with JOBS_LOCK:
        JOBS[job_id] = job
    return job


def get_job(job_id: str):
    return JOBS.get(job_id)


def run_job(job_id: str, excel_bytes: bytes, pdf_paths: list):
    """pdf_paths: list of (original_filename, path_on_disk). Runs in a background thread."""
    job = JOBS[job_id]
    try:
        try:
            records, warnings, total_rows, _ = load_excel(excel_bytes)
        except ExcelParseError as e:
            job.excel_error = str(e)
            return

        job.excel_warnings = warnings
        job.excel_total_rows = total_rows

        used_names = {}
        any_success = any_failed = False
        success_zip = zipfile.ZipFile(job.success_zip_path, "w", zipfile.ZIP_STORED)
        failed_zip = zipfile.ZipFile(job.failed_zip_path, "w", zipfile.ZIP_STORED)

        def fail(filename, content, reason):
            nonlocal any_failed
            job.failed.append({"original": filename, "reason": reason})
            failed_zip.writestr(filename, content)
            any_failed = True

        for filename, path in pdf_paths:
            job.current_filename = filename
            with open(path, "rb") as fh:
                content = fh.read()

            text, read_error = read_pdf_text(content)
            if read_error:
                fail(filename, content, read_error)
                job.processed += 1
                os.remove(path)
                continue

            filename_doc_no = extract_doc_no_from_filename(filename)
            doc_no = None

            if filename_doc_no and filename_doc_no in records and filename_doc_no in extract_all_numbers(text):
                doc_no = filename_doc_no
            else:
                content_matches = find_doc_no_matches_in_content(text, records)
                if len(content_matches) == 1:
                    doc_no = content_matches[0]
                elif len(content_matches) > 1:
                    fail(filename, content,
                         f"Multiple possible Doc. Nos found inside the PDF matching Excel entries: "
                         f"{', '.join(content_matches)}. Cannot determine automatically.")
                    job.processed += 1
                    os.remove(path)
                    continue
                else:
                    found_numbers = extract_all_numbers(text)
                    if filename_doc_no and filename_doc_no in records:
                        reason = f"Doc. No. '{filename_doc_no}' matched via filename but not found in PDF content."
                    elif found_numbers:
                        reason = f"Found number(s) inside PDF ({', '.join(sorted(found_numbers))}) not in Excel."
                    elif filename_doc_no:
                        reason = f"Doc. No. '{filename_doc_no}' from filename not in Excel, none found in PDF."
                    else:
                        reason = "No number matching an Excel Doc. No. was found inside the PDF."
                    fail(filename, content, reason)
                    job.processed += 1
                    os.remove(path)
                    continue

            record = records[doc_no]
            record["matched"] = True
            new_name = build_new_filename(record, doc_no)
            final_name = new_name
            if final_name in used_names:
                used_names[final_name] += 1
                stem, ext = final_name.rsplit(".", 1)
                final_name = f"{stem}__dup{used_names[new_name]}.{ext}"
            else:
                used_names[final_name] = 0

            job.success.append({"original": filename, "renamed": final_name})
            success_zip.writestr(final_name, content)
            any_success = True
            job.processed += 1
            os.remove(path)  # free disk immediately, don't wait till the end

        success_zip.close()
        failed_zip.close()
        if not any_success:
            os.remove(job.success_zip_path)
        if not any_failed:
            os.remove(job.failed_zip_path)

        for doc_no, rec in records.items():
            if not rec["matched"]:
                job.excel_unmatched.append({
                    "doc_no": doc_no, "document_type": rec["document_type"],
                    "customer": rec["customer"], "posting_date": rec["posting_date"],
                })
        job.excel_matched_count = len(records) - len(job.excel_unmatched)

    except Exception as e:
        job.error = f"{type(e).__name__}: {e}"
    finally:
        job.done = True