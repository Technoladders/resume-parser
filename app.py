from flask import Flask, request, jsonify, make_response # <--- CHANGE HERE: Add make_response
import uuid
from flask_cors import CORS
from rq.job import Job
import tasks
from config import queue, redis_conn, supabase_admin_client as supabase
import logging

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.route('/api/validate-candidate', methods=['POST'])
def validate_candidate():
    # START OF CHANGE: Add this block of code here to fix the timeout.
    # This handles the browser's preflight check before it tries to read the request body.
    # END OF CHANGE
    logger.info("Received request to /api/validate-candidate with headers: %s", request.headers)
    data = request.get_json()
    job_id = data.get('job_id')
    candidate_id = data.get('candidate_id')
    resume_url = data.get('resume_url')
    job_description = data.get('job_description')
    organization_id = data.get('organization_id')
    user_id = data.get('user_id')

    if not all([job_id, candidate_id, resume_url, job_description, organization_id, user_id ]):
        logger.error("Missing required fields: %s", data)
        return jsonify({"error": "Missing required fields"}), 400

    # Validate job_id exists in hr_jobs
    try:
        response = supabase.table("hr_jobs").select("id, job_id").eq("job_id", job_id).eq("organization_id", organization_id).execute()
        logger.info("Supabase response for job_id %s: %s", job_id, response)
        if not response.data:
            all_jobs = supabase.table("hr_jobs").select("job_id").execute()
            logger.info("All job IDs in hr_jobs: %s", all_jobs.data)
            logger.error("Job ID %s not found in hr_jobs for organization_id", job_id, organization_id)
            return jsonify({"error": "Job ID not found for the specified organization"}), 404
        job_uuid = response.data[0]["id"]
    except Exception as e:
        logger.error("Error validating job_id %s for organization_id  %s: %s", job_id, organization_id, str(e))
        return jsonify({"error": "Failed to validate job ID"}), 500

    job = queue.enqueue(tasks.process_analysis, job_uuid, candidate_id, resume_url, job_description, organization_id, user_id)
    logger.info("Enqueued job with ID: %s", job.id)
    return jsonify({"job_id": job.id, "job_uuid": job_uuid}), 202

# --- NEW BATCH PROCESSING ENDPOINT ---

@app.route('/api/validate-candidates-batch', methods=['POST'])
def validate_candidates_batch():
    # START OF CHANGE: Add this block of code here to fix the timeout.
    # This handles the browser's preflight check before it tries to read the request body.
    # END OF CHANGE
    """
    Accepts a batch of candidates for a single job and queues them for analysis.
    """
    logger.info("Received request to /api/validate-candidates-batch")
    data = request.get_json()

    # --- 1. Get common data for the whole batch ---
    job_id = data.get('job_id')
    job_description = data.get('job_description')
    organization_id = data.get('organization_id')
    user_id = data.get('user_id')
    candidates = data.get('candidates') # This should be a list of candidate objects

    if not all([job_id, job_description, organization_id, user_id, candidates]):
        return jsonify({"error": "Missing job_id, job_description, organization_id, user_id, or candidates list"}), 400
    
    if not isinstance(candidates, list) or not candidates:
        return jsonify({"error": "'candidates' must be a non-empty list"}), 400

    # --- 2. Validate the Job ID once for the entire batch ---
    try:
        response = supabase.table("hr_jobs").select("id").eq("job_id", job_id).eq("organization_id", organization_id).execute()
        if not response.data:
            logger.error(f"Batch job failed: Job ID {job_id} not found for organization {organization_id}")
            return jsonify({"error": "Job ID not found for the specified organization"}), 404
        job_uuid = response.data[0]["id"]
    except Exception as e:
        logger.error(f"Batch job failed: Error validating job_id {job_id}: {e}")
        return jsonify({"error": "Failed to validate job ID"}), 500

    # --- 3. Loop through candidates and enqueue a job for each one ---
    enqueued_jobs = []
    batch_id = str(uuid.uuid4()) # A unique ID for this entire batch
    
    for candidate in candidates:
        candidate_id = candidate.get('candidate_id')
        resume_url = candidate.get('resume_url')

        if not all([candidate_id, resume_url]):
            logger.warning(f"Skipping candidate in batch {batch_id} due to missing data: {candidate}")
            continue # Skip this invalid entry and continue with the rest
        
        # Enqueue the background task from tasks.py
        job = queue.enqueue(
            tasks.process_analysis,
            job_uuid,
            candidate_id,
            resume_url,
            job_description,
            organization_id,
            user_id
        )
        logger.info(f"Enqueued job {job.id} for candidate {candidate_id} in batch {batch_id}")
        enqueued_jobs.append({
            "candidate_id": candidate_id,
            "job_id": job.id # This is the RQ job ID for tracking
        })

    # --- 4. Return the list of enqueued jobs ---
    return jsonify({
        "message": f"Successfully enqueued {len(enqueued_jobs)} candidates for processing.",
        "batch_id": batch_id,
        "job_uuid": job_uuid,
        "enqueued_jobs": enqueued_jobs
    }), 202

@app.route('/api/job-status/<job_id>', methods=['GET'])
def job_status(job_id):
    try:
        job = Job.fetch(job_id, connection=redis_conn)
        return jsonify({"status": job.get_status(), "result": job.result})
    except Exception as e:
        return jsonify({"error": str(e)}), 404

@app.route('/api/job-logs/<job_id>', methods=['GET'])
def job_logs(job_id):
    logs = redis_conn.lrange(f"job_logs:{job_id}", 0, -1)
    logs = [json.loads(log.decode('utf-8')) for log in logs]
    return jsonify({"logs": logs})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5005)