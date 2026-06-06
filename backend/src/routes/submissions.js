const express = require('express');
const router = express.Router();
const submissionController = require('../controllers/submissionController');

router.get('/', submissionController.getSubmissions);
router.post('/', submissionController.createSubmission);
router.get('/:id/download', submissionController.downloadSubmission);
router.get('/:id/files', submissionController.getSubmissionFiles);
router.get('/:id/download-file', submissionController.downloadSubmissionFile);

module.exports = router;
