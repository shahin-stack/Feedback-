/**
 * FeedbackIQ — Two Independent Report Workflows
 * Report 1: Feedback Analytics (4-sheet) → POST /process
 * Report 2: Monthly Branch Conversion     → POST /process-monthly
 */

document.addEventListener('DOMContentLoaded', () => {

    // =========================================================================
    // SHARED UTILITIES
    // =========================================================================
    function setNavStatus(text, state) {
        const nav = document.getElementById('navStatus');
        if (!nav) return;
        nav.querySelector('span').textContent = text;
        nav.querySelector('.status-dot').style.background =
            state === 'processing' ? '#f59e0b' :
            state === 'done'       ? '#10b981' : '#22d3ee';
    }

    function animateSteps(stepIds, intervalMs = 600) {
        return new Promise(resolve => {
            let i = 0;
            const tick = () => {
                if (i >= stepIds.length) { resolve(); return; }
                const el = document.getElementById(stepIds[i]);
                if (el) el.classList.add('done');
                i++;
                setTimeout(tick, intervalMs);
            };
            tick();
        });
    }

    function showPanel(panels, activeId) {
        panels.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.classList.add('hidden');
        });
        const active = document.getElementById(activeId);
        if (active) active.classList.remove('hidden');
    }

    function resetSteps(stepIds) {
        stepIds.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.classList.remove('done');
        });
    }

    // =========================================================================
    // REPORT 1 — FEEDBACK ANALYTICS
    // =========================================================================
    (function initReport1() {
        const salesInput   = document.getElementById('sales_file');
        const fbInput      = document.getElementById('feedback_file');
        const salesDZ      = document.getElementById('salesDropZone1');
        const fbDZ         = document.getElementById('fbDropZone1');
        const salesDC      = document.getElementById('salesDropContent1');
        const fbDC         = document.getElementById('fbDropContent1');
        const salesSC      = document.getElementById('salesSuccessContent1');
        const fbSC         = document.getElementById('fbSuccessContent1');
        const salesFN      = document.getElementById('salesFileName1');
        const fbFN         = document.getElementById('fbFileName1');
        const removeSales  = document.getElementById('removeSalesFile1');
        const removeFb     = document.getElementById('removeFbFile1');
        const generateBtn  = document.getElementById('generateBtn1');
        const generateInfo = document.getElementById('generateInfo1');
        const ushBadge     = document.getElementById('ushBadge1');
        const ushText      = document.getElementById('ushBadgeText1');
        const downloadBtn  = document.getElementById('downloadBtn1');
        const resetBtn     = document.getElementById('resetBtn1');
        const retryBtn     = document.getElementById('retryBtn1');
        const errMsg       = document.getElementById('errorMessage1');

        const PANELS    = ['uploadPanel1','processPanel1','successPanel1','errorPanel1'];
        const STEPS     = ['pstep1a','pstep2a','pstep3a','pstep4a'];
        const ENDPOINT  = '/process';

        function updateBadge() {
            const n = (salesInput.files.length > 0 ? 1 : 0) + (fbInput.files.length > 0 ? 1 : 0);
            const ready = n === 2;
            ushText.textContent = ready ? '✓ All files ready' : `${n} / 2 uploaded`;
            ushBadge.classList.toggle('all-ready', ready);
        }

        function checkReady() {
            const ready = salesInput.files.length > 0 && fbInput.files.length > 0;
            generateBtn.disabled = !ready;
            generateInfo.style.display = ready ? 'none' : 'flex';
            updateBadge();
        }

        function setupDropZone(dz, input, dropContent, successContent, fileNameEl) {
            dz.addEventListener('click', () => input.click());
            dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
            dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
            dz.addEventListener('drop', e => {
                e.preventDefault();
                dz.classList.remove('dragover');
                const file = e.dataTransfer.files[0];
                if (file) { setFile(input, dz, dropContent, successContent, fileNameEl, file); }
            });
            input.addEventListener('change', () => {
                if (input.files[0]) setFile(input, dz, dropContent, successContent, fileNameEl, input.files[0]);
            });
        }

        function setFile(input, dz, dc, sc, fn, file) {
            const dt = new DataTransfer();
            dt.items.add(file);
            input.files = dt.files;
            fn.textContent = file.name;
            dc.classList.add('hidden');
            sc.classList.remove('hidden');
            dz.classList.add('has-file');
            checkReady();
        }

        function clearFile(input, dz, dc, sc) {
            input.value = '';
            dc.classList.remove('hidden');
            sc.classList.add('hidden');
            dz.classList.remove('has-file');
            checkReady();
        }

        setupDropZone(salesDZ, salesInput, salesDC, salesSC, salesFN);
        setupDropZone(fbDZ, fbInput, fbDC, fbSC, fbFN);
        removeSales.addEventListener('click', e => { e.stopPropagation(); clearFile(salesInput, salesDZ, salesDC, salesSC); });
        removeFb.addEventListener('click',    e => { e.stopPropagation(); clearFile(fbInput, fbDZ, fbDC, fbSC); });

        generateBtn.addEventListener('click', async () => {
            showPanel(PANELS, 'processPanel1');
            setNavStatus('Generating Analytics...', 'processing');
            resetSteps(STEPS);

            const formData = new FormData();
            formData.append('sales_file', salesInput.files[0]);
            formData.append('feedback_file', fbInput.files[0]);

            const stepsPromise = animateSteps(STEPS, 700);

            try {
                const res  = await fetch(ENDPOINT, { method: 'POST', body: formData });
                const text = await res.text();
                let data;
                try { data = JSON.parse(text); }
                catch (_) {
                    throw new Error(
                        text.trim() === ''
                        ? 'Server timed out or ran out of memory. Please try with a smaller file.'
                        : 'Server returned an unexpected response. Check Render logs.'
                    );
                }
                await stepsPromise;

                if (data.status === 'success') {
                    downloadBtn.href = `/download/${data.filename}`;
                    downloadBtn.download = data.filename;
                    showPanel(PANELS, 'successPanel1');
                    setNavStatus('Report 1 Ready', 'done');
                } else {
                    errMsg.textContent = data.message || 'Unknown error';
                    showPanel(PANELS, 'errorPanel1');
                    setNavStatus('Error in Report 1', 'ready');
                }
            } catch (err) {
                await stepsPromise;
                errMsg.textContent = err.message;
                showPanel(PANELS, 'errorPanel1');
                setNavStatus('Error in Report 1', 'ready');
            }
        });

        resetBtn.addEventListener('click', () => {
            clearFile(salesInput, salesDZ, salesDC, salesSC);
            clearFile(fbInput, fbDZ, fbDC, fbSC);
            resetSteps(STEPS);
            showPanel(PANELS, 'uploadPanel1');
            setNavStatus('Ready to Process', 'ready');
        });

        retryBtn.addEventListener('click', () => {
            resetSteps(STEPS);
            showPanel(PANELS, 'uploadPanel1');
        });

        checkReady();
    })();


    // =========================================================================
    // REPORT 2 — MONTHLY BRANCH CONVERSION
    // =========================================================================
    (function initReport2() {
        const salesInput   = document.getElementById('sales_file_m');
        const fbInput      = document.getElementById('feedback_file_m');
        const salesDZ      = document.getElementById('salesDropZone2');
        const fbDZ         = document.getElementById('fbDropZone2');
        const salesDC      = document.getElementById('salesDropContent2');
        const fbDC         = document.getElementById('fbDropContent2');
        const salesSC      = document.getElementById('salesSuccessContent2');
        const fbSC         = document.getElementById('fbSuccessContent2');
        const salesFN      = document.getElementById('salesFileName2');
        const fbFN         = document.getElementById('fbFileName2');
        const removeSales  = document.getElementById('removeSalesFile2');
        const removeFb     = document.getElementById('removeFbFile2');
        const generateBtn  = document.getElementById('generateBtn2');
        const generateInfo = document.getElementById('generateInfo2');
        const ushBadge     = document.getElementById('ushBadge2');
        const ushText      = document.getElementById('ushBadgeText2');
        const downloadBtn  = document.getElementById('downloadBtn2');
        const resetBtn     = document.getElementById('resetBtn2');
        const retryBtn     = document.getElementById('retryBtn2');
        const errMsg       = document.getElementById('errorMessage2');

        const PANELS   = ['uploadPanel2','processPanel2','successPanel2','errorPanel2'];
        const STEPS    = ['pstep1b','pstep2b','pstep3b','pstep4b'];
        const ENDPOINT = '/process-monthly';

        function updateBadge() {
            const n = (salesInput.files.length > 0 ? 1 : 0) + (fbInput.files.length > 0 ? 1 : 0);
            const ready = n === 2;
            ushText.textContent = ready ? '✓ All files ready' : `${n} / 2 uploaded`;
            ushBadge.classList.toggle('all-ready', ready);
        }

        function checkReady() {
            const ready = salesInput.files.length > 0 && fbInput.files.length > 0;
            generateBtn.disabled = !ready;
            generateInfo.style.display = ready ? 'none' : 'flex';
            updateBadge();
        }

        function setupDropZone(dz, input, dropContent, successContent, fileNameEl) {
            dz.addEventListener('click', () => input.click());
            dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
            dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
            dz.addEventListener('drop', e => {
                e.preventDefault();
                dz.classList.remove('dragover');
                const file = e.dataTransfer.files[0];
                if (file) setFile(input, dz, dropContent, successContent, fileNameEl, file);
            });
            input.addEventListener('change', () => {
                if (input.files[0]) setFile(input, dz, dropContent, successContent, fileNameEl, input.files[0]);
            });
        }

        function setFile(input, dz, dc, sc, fn, file) {
            const dt = new DataTransfer();
            dt.items.add(file);
            input.files = dt.files;
            fn.textContent = file.name;
            dc.classList.add('hidden');
            sc.classList.remove('hidden');
            dz.classList.add('has-file');
            checkReady();
        }

        function clearFile(input, dz, dc, sc) {
            input.value = '';
            dc.classList.remove('hidden');
            sc.classList.add('hidden');
            dz.classList.remove('has-file');
            checkReady();
        }

        setupDropZone(salesDZ, salesInput, salesDC, salesSC, salesFN);
        setupDropZone(fbDZ, fbInput, fbDC, fbSC, fbFN);
        removeSales.addEventListener('click', e => { e.stopPropagation(); clearFile(salesInput, salesDZ, salesDC, salesSC); });
        removeFb.addEventListener('click',    e => { e.stopPropagation(); clearFile(fbInput, fbDZ, fbDC, fbSC); });

        generateBtn.addEventListener('click', async () => {
            showPanel(PANELS, 'processPanel2');
            setNavStatus('Generating SMS Report...', 'processing');
            resetSteps(STEPS);

            const formData = new FormData();
            formData.append('sales_file_m', salesInput.files[0]);
            formData.append('feedback_file_m', fbInput.files[0]);

            const stepsPromise = animateSteps(STEPS, 700);

            try {
                const res  = await fetch(ENDPOINT, { method: 'POST', body: formData });
                const text = await res.text();
                let data;
                try { data = JSON.parse(text); }
                catch (_) {
                    throw new Error(
                        text.trim() === ''
                        ? 'Server timed out or ran out of memory. Please try with a smaller file.'
                        : 'Server returned an unexpected response. Check Render logs.'
                    );
                }
                await stepsPromise;

                if (data.status === 'success') {
                    downloadBtn.href = `/download/${data.filename}`;
                    downloadBtn.download = data.filename;
                    showPanel(PANELS, 'successPanel2');
                    setNavStatus('Report 2 Ready', 'done');
                } else {
                    errMsg.textContent = data.message || 'Unknown error';
                    showPanel(PANELS, 'errorPanel2');
                    setNavStatus('Error in Report 2', 'ready');
                }
            } catch (err) {
                await stepsPromise;
                errMsg.textContent = err.message;
                showPanel(PANELS, 'errorPanel2');
                setNavStatus('Error in Report 2', 'ready');
            }
        });

        resetBtn.addEventListener('click', () => {
            clearFile(salesInput, salesDZ, salesDC, salesSC);
            clearFile(fbInput, fbDZ, fbDC, fbSC);
            resetSteps(STEPS);
            showPanel(PANELS, 'uploadPanel2');
            setNavStatus('Ready to Process', 'ready');
        });

        retryBtn.addEventListener('click', () => {
            resetSteps(STEPS);
            showPanel(PANELS, 'uploadPanel2');
        });

        checkReady();
    })();

});
