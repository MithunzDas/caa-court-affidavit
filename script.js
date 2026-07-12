/* ============================================================
   CAA Affidavit Generator — Application Logic
   Handles: View navigation, form validation, API submission,
            and download link generation.
   ============================================================ */

(function () {
    'use strict';

    // ──────────────────────────────────────────────────────────
    // Configuration
    // ──────────────────────────────────────────────────────────

    /**
     * n8n webhook URL.
     * Replace this with your actual n8n webhook endpoint after deployment.
     * Example: 'https://n8n.yourdomain.com/webhook/generate-affidavit'
     */
    const API_URL = 'https://n8n.srv1594654.hstgr.cloud/webhook/generate-affidavit';

    /**
     * List of all form field names in submission order.
     * Used for iteration during validation and data collection.
     */
    const FIELD_NAMES = [
        'applicant_name', 'guardian_type', 'father_name',
        'india_village', 'india_po', 'india_ps',
        'india_district', 'india_pin', 'india_state',
        'bd_village', 'bd_po', 'bd_ps', 'bd_district',
        'entry_date', 'verification_date', 'advocate',
        'witness_name', 'witness_guardian_type', 'witness_father',
        'witness_age', 'witness_occupation',
        'witness_village', 'witness_po', 'witness_ps',
        'witness_district', 'witness_pin', 'witness_state'
    ];

    /**
     * Validation rules per field.
     * Each key maps to a function that returns an error message string,
     * or an empty string if valid.
     */
    const VALIDATORS = {
        applicant_name:       (v) => minLen(v, 2, 'Full name'),
        guardian_type:        (v) => required(v, 'Relation'),
        father_name:          (v) => minLen(v, 2, 'Guardian name'),
        india_village:        (v) => minLen(v, 2, 'Village'),
        india_po:             (v) => minLen(v, 2, 'Post office'),
        india_ps:             (v) => minLen(v, 2, 'Police station'),
        india_district:       (v) => minLen(v, 2, 'District'),
        india_pin:            (v) => pin(v),
        india_state:          (v) => minLen(v, 2, 'State'),
        bd_village:           (v) => minLen(v, 2, 'Village'),
        bd_po:                (v) => minLen(v, 2, 'Post office'),
        bd_ps:                (v) => minLen(v, 2, 'Police station'),
        bd_district:          (v) => minLen(v, 2, 'District'),
        entry_date:           (v) => dateFormat(v, 'Entry date'),
        verification_date:    (v) => dateFormat(v, 'Verification date'),
        advocate:             (v) => required(v, 'Advocate'),
        witness_name:         (v) => minLen(v, 2, 'Witness name'),
        witness_guardian_type:(v) => required(v, 'Relation'),
        witness_father:       (v) => minLen(v, 2, 'Guardian name'),
        witness_age:          (v) => age(v),
        witness_occupation:   (v) => minLen(v, 2, 'Occupation'),
        witness_village:      (v) => minLen(v, 2, 'Village'),
        witness_po:           (v) => minLen(v, 2, 'Post office'),
        witness_ps:           (v) => minLen(v, 2, 'Police station'),
        witness_district:     (v) => minLen(v, 2, 'District'),
        witness_pin:          (v) => pin(v),
        witness_state:        (v) => minLen(v, 2, 'State')
    };


    // ──────────────────────────────────────────────────────────
    // Validation Helper Functions
    // ──────────────────────────────────────────────────────────

    /** Check for non-empty value */
    function required(value, label) {
        return value.trim() === '' ? `${label} is required` : '';
    }

    /** Check minimum string length */
    function minLen(value, min, label) {
        const v = value.trim();
        if (v === '') return `${label} is required`;
        if (v.length < min) return `${label} must be at least ${min} characters`;
        return '';
    }

    /** Validate 6-digit PIN code */
    function pin(value) {
        const v = value.trim();
        if (v === '') return 'PIN code is required';
        if (!/^\d{6}$/.test(v)) return 'PIN code must be exactly 6 digits';
        return '';
    }

    /** Validate DD/MM/YYYY date format */
    function dateFormat(value, label) {
        const v = value.trim();
        if (v === '') return `${label} is required`;
        if (!/^\d{2}\/\d{2}\/\d{4}$/.test(v)) return `${label} must be in DD/MM/YYYY format`;

        // Basic range check for day/month
        const parts = v.split('/');
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);

        if (month < 1 || month > 12) return 'Invalid month';
        if (day < 1 || day > 31) return 'Invalid day';
        if (year < 1900 || year > 2100) return 'Invalid year';

        return '';
    }

    /** Validate age (1-3 digits, reasonable range) */
    function age(value) {
        const v = value.trim();
        if (v === '') return 'Age is required';
        if (!/^\d{1,3}$/.test(v)) return 'Age must be a number';
        const num = parseInt(v, 10);
        if (num < 1 || num > 150) return 'Age must be between 1 and 150';
        return '';
    }


    // ──────────────────────────────────────────────────────────
    // DOM References
    // ──────────────────────────────────────────────────────────

    const viewDashboard = document.getElementById('view-dashboard');
    const viewForm      = document.getElementById('view-form');
    const viewSuccess   = document.getElementById('view-success');

    const btnCreate          = document.getElementById('btn-create');
    const btnBack            = document.getElementById('btn-back');
    const btnReset           = document.getElementById('btn-reset');
    const btnGenerate        = document.getElementById('btn-generate');
    const btnGenerateAnother = document.getElementById('btn-generate-another');

    const form              = document.getElementById('affidavit-form');
    const successFilename   = document.getElementById('success-filename');
    const btnDownloadDocx   = document.getElementById('btn-download-docx');
    const btnDownloadPdf    = document.getElementById('btn-download-pdf');


    // ──────────────────────────────────────────────────────────
    // View Navigation
    // ──────────────────────────────────────────────────────────

    /**
     * Switch the active view.
     * Only one view is visible at a time.
     * @param {'dashboard' | 'form' | 'success'} viewName
     */
    function showView(viewName) {
        viewDashboard.classList.remove('active');
        viewForm.classList.remove('active');
        viewSuccess.classList.remove('active');

        switch (viewName) {
            case 'dashboard':
                viewDashboard.classList.add('active');
                break;
            case 'form':
                viewForm.classList.add('active');
                break;
            case 'success':
                viewSuccess.classList.add('active');
                break;
        }

        // Scroll to top on view change
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }


    // ──────────────────────────────────────────────────────────
    // Auto-format Date Inputs
    // ──────────────────────────────────────────────────────────

    /**
     * Automatically insert '/' separators as the user types a date.
     * Triggers on the entry_date and verification_date fields.
     */
    function autoFormatDate(event) {
        const input = event.target;
        let value = input.value.replace(/[^\d]/g, ''); // strip non-digits

        if (value.length > 8) {
            value = value.substring(0, 8);
        }

        // Insert slashes: DD/MM/YYYY
        if (value.length > 4) {
            value = value.substring(0, 2) + '/' + value.substring(2, 4) + '/' + value.substring(4);
        } else if (value.length > 2) {
            value = value.substring(0, 2) + '/' + value.substring(2);
        }

        input.value = value;
    }

    // Attach date formatters
    document.getElementById('entry_date').addEventListener('input', autoFormatDate);
    document.getElementById('verification_date').addEventListener('input', autoFormatDate);


    // ──────────────────────────────────────────────────────────
    // Auto-fill Verification Date with Today
    // ──────────────────────────────────────────────────────────

    /**
     * Set the verification_date field to today's date (DD/MM/YYYY).
     * Called when the form view opens.
     */
    function setTodayAsVerificationDate() {
        const today = new Date();
        const dd = String(today.getDate()).padStart(2, '0');
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const yyyy = today.getFullYear();
        document.getElementById('verification_date').value = `${dd}/${mm}/${yyyy}`;
    }


    // ──────────────────────────────────────────────────────────
    // Form Validation
    // ──────────────────────────────────────────────────────────

    /**
     * Validate a single field and show/hide its error message.
     * @param {string} fieldName
     * @returns {boolean} true if valid
     */
    function validateField(fieldName) {
        const input = document.getElementById(fieldName);
        const errorEl = document.getElementById(`error-${fieldName}`);
        const group = input.closest('.form-group');
        const validator = VALIDATORS[fieldName];

        if (!validator) return true;

        const errorMsg = validator(input.value);

        if (errorMsg) {
            group.classList.add('has-error');
            errorEl.textContent = errorMsg;
            return false;
        } else {
            group.classList.remove('has-error');
            errorEl.textContent = '';
            return true;
        }
    }

    /**
     * Validate all fields. Returns true if the entire form is valid.
     * Focuses the first invalid field.
     */
    function validateForm() {
        let isValid = true;
        let firstInvalid = null;

        for (const fieldName of FIELD_NAMES) {
            const fieldValid = validateField(fieldName);
            if (!fieldValid && isValid) {
                isValid = false;
                firstInvalid = document.getElementById(fieldName);
            }
        }

        // Focus and scroll to the first invalid field
        if (firstInvalid) {
            firstInvalid.focus();
            firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }

        return isValid;
    }

    /**
     * Attach real-time validation: clear error on input change.
     * This provides instant feedback as the user corrects a field.
     */
    function attachRealtimeValidation() {
        for (const fieldName of FIELD_NAMES) {
            const input = document.getElementById(fieldName);
            const eventType = input.tagName === 'SELECT' ? 'change' : 'input';

            input.addEventListener(eventType, function () {
                // Only re-validate if the field currently has an error
                const group = input.closest('.form-group');
                if (group.classList.contains('has-error')) {
                    validateField(fieldName);
                }
            });
        }
    }


    // ──────────────────────────────────────────────────────────
    // Collect Form Data
    // ──────────────────────────────────────────────────────────

    /**
     * Collect all form field values into a plain object.
     * Text fields are trimmed. Values are sent as-is (uppercase is
     * expected to be typed by the user per legal convention).
     * @returns {Object}
     */
    function collectFormData() {
        const data = {};
        for (const fieldName of FIELD_NAMES) {
            const input = document.getElementById(fieldName);
            data[fieldName] = input.value.trim();
        }
        return data;
    }


    // ──────────────────────────────────────────────────────────
    // API Submission
    // ──────────────────────────────────────────────────────────

    /**
     * Submit form data to the n8n webhook endpoint.
     * Handles loading state, success, and error responses.
     * @param {Object} data - Collected form data
     */
    async function submitForm(data) {
        // Enter loading state
        btnGenerate.classList.add('loading');
        btnGenerate.disabled = true;

        try {
            const response = await fetch(API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(data)
            });

            if (!response.ok) {
                const errorBody = await response.text();
                throw new Error(`Server error (${response.status}): ${errorBody}`);
            }

            const result = await response.json();

            if (result.success) {
                // Display success view with download links
                successFilename.textContent = `${data.applicant_name.replace(/\s+/g, '_')}`;
                btnDownloadDocx.href = result.docx;
                btnDownloadPdf.href = result.pdf;
                showView('success');
            } else {
                throw new Error(result.message || 'Document generation failed');
            }
        } catch (error) {
            console.error('Submission error:', error);

            // Show user-friendly error
            showError(error.message || 'Failed to connect to the server. Please try again.');
        } finally {
            // Exit loading state
            btnGenerate.classList.remove('loading');
            btnGenerate.disabled = false;
        }
    }


    // ──────────────────────────────────────────────────────────
    // Error Display
    // ──────────────────────────────────────────────────────────

    /**
     * Display a temporary error notification at the top of the form.
     * Auto-dismisses after 8 seconds or on click.
     * @param {string} message
     */
    function showError(message) {
        // Remove existing error banner if any
        const existing = document.querySelector('.error-banner');
        if (existing) existing.remove();

        const banner = document.createElement('div');
        banner.className = 'error-banner';
        banner.innerHTML = `
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="20" height="20">
                <circle cx="12" cy="12" r="10"/>
                <line x1="15" y1="9" x2="9" y2="15"/>
                <line x1="9" y1="9" x2="15" y2="15"/>
            </svg>
            <span>${escapeHtml(message)}</span>
            <button type="button" class="error-dismiss" aria-label="Dismiss">&times;</button>
        `;

        // Style the banner (injected inline to keep CSS file clean)
        Object.assign(banner.style, {
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            padding: '14px 18px',
            marginBottom: '16px',
            background: 'hsla(0, 70%, 55%, 0.12)',
            border: '1px solid hsla(0, 70%, 55%, 0.3)',
            borderRadius: '10px',
            color: 'hsl(0, 70%, 70%)',
            fontSize: '0.875rem',
            animation: 'fadeIn 250ms ease'
        });

        // Insert at top of form
        form.insertBefore(banner, form.firstChild);

        // Dismiss handlers
        const dismiss = () => {
            banner.style.opacity = '0';
            banner.style.transform = 'translateY(-8px)';
            banner.style.transition = 'opacity 200ms, transform 200ms';
            setTimeout(() => banner.remove(), 200);
        };

        banner.querySelector('.error-dismiss').addEventListener('click', dismiss);
        setTimeout(dismiss, 8000);
    }

    /**
     * Basic HTML escaping to prevent XSS in error messages.
     * @param {string} str
     * @returns {string}
     */
    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }


    // ──────────────────────────────────────────────────────────
    // Reset Form
    // ──────────────────────────────────────────────────────────

    /**
     * Clear all form fields and remove error states.
     */
    function resetForm() {
        form.reset();

        // Clear all error states
        const errorGroups = form.querySelectorAll('.form-group.has-error');
        errorGroups.forEach(group => group.classList.remove('has-error'));

        const errorMessages = form.querySelectorAll('.error-message');
        errorMessages.forEach(el => { el.textContent = ''; });

        // Remove any error banners
        const banner = form.querySelector('.error-banner');
        if (banner) banner.remove();

        // Re-set verification date to today
        setTodayAsVerificationDate();
    }


    // ──────────────────────────────────────────────────────────
    // Event Listeners
    // ──────────────────────────────────────────────────────────

    // Dashboard → Form
    btnCreate.addEventListener('click', function () {
        resetForm();
        showView('form');
    });

    // Form → Dashboard
    btnBack.addEventListener('click', function () {
        showView('dashboard');
    });

    // Reset button
    btnReset.addEventListener('click', function () {
        resetForm();
        // Scroll to top of form
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    // Form submission
    form.addEventListener('submit', function (event) {
        event.preventDefault();

        if (!validateForm()) {
            return;
        }

        const data = collectFormData();
        submitForm(data);
    });

    // Success → Form (generate another)
    btnGenerateAnother.addEventListener('click', function () {
        resetForm();
        showView('form');
    });

    // Initialize real-time validation listeners
    attachRealtimeValidation();

})();
