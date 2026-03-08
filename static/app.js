/* ── Toggle password visibility ─────────────────── */
function togglePwd(id) {
  const el = document.getElementById(id);
  el.type = (el.type === 'password') ? 'text' : 'password';
}

/* ── Password strength meter ─────────────────────── */
const pwdInput = document.getElementById('new_password');
if (pwdInput) {
  pwdInput.addEventListener('input', function () {
    const val = this.value;
    let score = 0;
    if (val.length >= 8)   score++;
    if (val.length >= 12)  score++;
    if (/[A-Z]/.test(val)) score++;
    if (/[0-9]/.test(val)) score++;
    if (/[^A-Za-z0-9]/.test(val)) score++;

    const bar   = document.getElementById('strengthBar');
    const label = document.getElementById('strengthLabel');
    const widths = ['0%','25%','50%','75%','90%','100%'];
    const colors = ['#f85149','#f85149','#d29922','#d29922','#3fb950','#3fb950'];
    const labels = ['','Very Weak','Weak','Fair','Strong','Very Strong'];

    bar.style.width      = widths[score];
    bar.style.background = colors[score];
    label.textContent    = labels[score];
    label.style.color    = colors[score];
  });
}

/* ── Drag & Drop Upload ──────────────────────────── */
const dropZone    = document.getElementById('dropZone');
const fileInput   = document.getElementById('fileInput');
const filePreview = document.getElementById('filePreview');
const fileName    = document.getElementById('fileName');
const fileSize    = document.getElementById('fileSize');
const clearFile   = document.getElementById('clearFile');
const generateBtn = document.getElementById('generateBtn');
const progressWrap = document.getElementById('progressWrap');
const progressBar  = document.getElementById('progressBar');
const resultCard  = document.getElementById('resultCard');
const errorCard   = document.getElementById('errorCard');
const downloadLink = document.getElementById('downloadLink');
const resetBtn    = document.getElementById('resetBtn');
const errorResetBtn = document.getElementById('errorResetBtn');
const errorMsg    = document.getElementById('errorMsg');

let selectedFile = null;

function formatBytes(bytes) {
  if (bytes < 1024)       return bytes + ' B';
  if (bytes < 1048576)    return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(2) + ' MB';
}

function showFile(file) {
  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = formatBytes(file.size);
  dropZone.classList.add('hidden');
  filePreview.classList.remove('hidden');
  resultCard.classList.add('hidden');
  errorCard.classList.add('hidden');
}

function resetUpload() {
  selectedFile = null;
  fileInput.value = '';
  dropZone.classList.remove('hidden');
  filePreview.classList.add('hidden');
  progressWrap.classList.add('hidden');
  resultCard.classList.add('hidden');
  errorCard.classList.add('hidden');
  progressBar.style.width = '0%';
  generateBtn.disabled = false;
  generateBtn.textContent = '⚡ Generate Report';
}

if (dropZone) {
  dropZone.addEventListener('click', () => fileInput.click());

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) showFile(file);
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) showFile(fileInput.files[0]);
  });

  clearFile.addEventListener('click', resetUpload);
  if (resetBtn)     resetBtn.addEventListener('click', resetUpload);
  if (errorResetBtn) errorResetBtn.addEventListener('click', resetUpload);

  generateBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    generateBtn.disabled    = true;
    generateBtn.textContent = '⏳ Processing…';
    progressWrap.classList.remove('hidden');

    // Animate progress bar while waiting
    let fakeProgress = 0;
    const ticker = setInterval(() => {
      fakeProgress = Math.min(fakeProgress + Math.random() * 8, 88);
      progressBar.style.width = fakeProgress + '%';
    }, 250);

    try {
      const formData = new FormData();
      formData.append('file', selectedFile);

      const response = await fetch('/upload', {
        method: 'POST',
        body: formData
      });

      clearInterval(ticker);

      if (!response.ok) {
        const json = await response.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(json.error || `Server error: ${response.status}`);
      }

      progressBar.style.width = '100%';
      await new Promise(r => setTimeout(r, 300));

      const blob = await response.blob();
      const url  = URL.createObjectURL(blob);

      // Derive filename from Content-Disposition or fallback
      let dlName = 'Report_Processed.xlsx';
      const cd = response.headers.get('Content-Disposition');
      if (cd) {
        const match = cd.match(/filename[^;=\n]*=['"]?([^'"\n]+)/i);
        if (match) dlName = match[1];
      }

      downloadLink.href = url;
      downloadLink.download = dlName;

      filePreview.classList.add('hidden');
      resultCard.classList.remove('hidden');

    } catch (err) {
      clearInterval(ticker);
      progressWrap.classList.add('hidden');
      filePreview.classList.add('hidden');
      errorCard.classList.remove('hidden');
      errorMsg.textContent = err.message || 'An unexpected error occurred.';
    }
  });
}

/* ── Auto-dismiss flashes ────────────────────────── */
document.querySelectorAll('.flash').forEach(el => {
  setTimeout(() => el.style.opacity = '0', 5000);
  setTimeout(() => el.remove(),            5400);
  el.style.transition = 'opacity 0.4s';
});
