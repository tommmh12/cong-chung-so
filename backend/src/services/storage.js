const { createClient } = require('@supabase/supabase-js');
const fs = require('fs');
const path = require('path');
const os = require('os');

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

const BUCKET = process.env.SUPABASE_STORAGE_BUCKET || 'templates';

// Upload local file → Supabase Storage, returns storagePath
async function uploadFile(localPath, storagePath) {
  const buffer = fs.readFileSync(localPath);
  const { error } = await supabase.storage
    .from(BUCKET)
    .upload(storagePath, buffer, {
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      upsert: true
    });
  if (error) throw new Error(`Supabase upload failed: ${error.message}`);
  return storagePath;
}

// Upload buffer → Supabase Storage
async function uploadBuffer(buffer, storagePath) {
  const { error } = await supabase.storage
    .from(BUCKET)
    .upload(storagePath, buffer, {
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      upsert: true
    });
  if (error) throw new Error(`Supabase upload failed: ${error.message}`);
  return storagePath;
}

// Download from Supabase → Buffer
async function downloadToBuffer(storagePath) {
  const { data, error } = await supabase.storage.from(BUCKET).download(storagePath);
  if (error) throw new Error(`Supabase download failed: ${error.message}`);
  return Buffer.from(await data.arrayBuffer());
}

// Download from Supabase → temp local file (auto-cleanup via returned fn)
async function downloadToTemp(storagePath) {
  const buffer = await downloadToBuffer(storagePath);
  const tmpPath = path.join(os.tmpdir(), `lx-${Date.now()}-${path.basename(storagePath)}`);
  fs.writeFileSync(tmpPath, buffer);
  const cleanup = () => { try { fs.unlinkSync(tmpPath); } catch (_) {} };
  return { tmpPath, cleanup };
}

// Delete from Supabase Storage
async function deleteFile(storagePath) {
  const { error } = await supabase.storage.from(BUCKET).remove([storagePath]);
  if (error) console.error('Storage delete error:', error.message);
}

// Returns true if this key is a Supabase path (not a legacy local path)
function isSupabasePath(storageKey) {
  return storageKey && !storageKey.startsWith('uploads/');
}

// Get a file as Buffer — handles both Supabase and legacy local paths
async function getFileBuffer(storageKey, backendRoot) {
  if (isSupabasePath(storageKey)) {
    return await downloadToBuffer(storageKey);
  }
  // Legacy: read from local disk
  const absPath = path.join(backendRoot, storageKey.replace(/\\/g, '/'));
  return fs.readFileSync(absPath);
}

// Call once on startup to ensure the bucket exists
async function ensureBucket() {
  const { error } = await supabase.storage.createBucket(BUCKET, { public: false });
  if (error && !error.message.includes('already exists')) {
    console.error('⚠️  Supabase bucket creation failed:', error.message);
  } else {
    console.log(`✅ Supabase Storage bucket "${BUCKET}" ready`);
  }
}

module.exports = {
  supabase,
  BUCKET,
  uploadFile,
  uploadBuffer,
  downloadToBuffer,
  downloadToTemp,
  deleteFile,
  isSupabasePath,
  getFileBuffer,
  ensureBucket
};
