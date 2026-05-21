document.addEventListener('DOMContentLoaded', () => {
    // Mobile Sidebar Toggle
    const mobileToggle = document.getElementById('mobile-toggle');
    const sidebar = document.getElementById('sidebar');
    
    if (mobileToggle && sidebar) {
        mobileToggle.addEventListener('click', () => {
            sidebar.classList.toggle('show');
        });

        // Close sidebar when clicking outside on mobile
        document.addEventListener('click', (e) => {
            if (window.innerWidth <= 768) {
                if (!sidebar.contains(e.target) && !mobileToggle.contains(e.target) && sidebar.classList.contains('show')) {
                    sidebar.classList.remove('show');
                }
            }
        });
    }

    // Course Search/Filter
    const courseSearch = document.getElementById('courseSearch');
    if (courseSearch) {
        courseSearch.addEventListener('keyup', (e) => {
            const term = e.target.value.toLowerCase();
            const cards = document.querySelectorAll('.course-card');
            
            cards.forEach(card => {
                const title = card.querySelector('.course-title').innerText.toLowerCase();
                const code = card.querySelector('.course-code-badge').innerText.toLowerCase();
                if (title.includes(term) || code.includes(term)) {
                    card.style.display = 'flex';
                } else {
                    card.style.display = 'none';
                }
            });
        });
    }

    // Timetable Generator Form
    const generateBtn = document.getElementById('btn-generate-timetable');
    if (generateBtn) {
        generateBtn.addEventListener('click', async () => {
            // Get values
            const semester = document.getElementById('semester').value;
            const maxPerDay = document.getElementById('max_per_day').value;
            const startHour = document.getElementById('start_hour').value;
            const endHour = document.getElementById('end_hour').value;

            // UI Loading state
            const originalText = generateBtn.innerHTML;
            generateBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating...';
            generateBtn.disabled = true;

            try {
                const response = await fetch('/admin/generate-timetable', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        semester: semester,
                        max_per_day: maxPerDay,
                        start_hour: startHour,
                        end_hour: endHour
                    })
                });

                const data = await response.json();
                
                if (data.success) {
                    showToast(`Success! Generated ${data.count} timetable entries.`, 'success');
                    
                    // Update stats
                    document.getElementById('stat-entries').innerText = data.count;
                    document.getElementById('stat-conflicts').innerText = data.conflicts;
                    
                    // Reload page after a short delay to show the new table
                    setTimeout(() => {
                        window.location.reload();
                    }, 1500);
                } else {
                    showToast('Failed to generate timetable.', 'error');
                }
            } catch (error) {
                console.error(error);
                showToast('An error occurred during generation.', 'error');
            } finally {
                generateBtn.innerHTML = originalText;
                generateBtn.disabled = false;
            }
        });
    }

    initializeExcelTimetableBuilder();
    initializeAdminTools();

    // i18n Translation Dictionary
    const translations = {
        'en': {
            'nav_dashboard': 'Dashboard',
            'nav_courses': 'Courses',
            'nav_grades': 'Grades',
            'nav_profile': 'Profile',
            'nav_admin_dashboard': 'Admin Dashboard',
            'nav_smart_timetable': 'Smart Timetable',
            'nav_students': 'Students',
            'nav_teachers': 'Teachers',
            'nav_rooms': 'Rooms',
            'nav_groups': 'Groups',
            'nav_generated_timetable': 'Generated Timetable',
            'nav_settings': 'Settings',
            'logout': 'Logout',
            'welcome': 'Welcome Back',
            'sign_in_hint': 'Sign in to your university account',
            'email_label': 'University Email',
            'password_label': 'Password',
            'sign_in_btn': 'Sign In',
            'demo_acc': 'Demo Accounts',
            'admin_dashboard_title': 'University Administration',
            'system_status': 'System Status',
            'open_builder': 'Open Timetable Builder',
            'total_students': 'Total Students',
            'total_teachers': 'Total Teachers',
            'conflicts': 'Conflicts',
            'quick_actions': 'Quick Actions',
            'generate_timetable': 'Generate Timetable',
            'manage_students': 'Manage Students',
            'manage_teachers': 'Manage Teachers',
            'manage_courses': 'Manage Courses',
            'manage_rooms': 'Manage Rooms',
            'recent_entries': 'Recently Generated Entries',
            'day': 'Day',
            'time': 'Time',
            'course': 'Course',
            'teacher': 'Teacher',
            'room': 'Room',
            'group': 'Group',
            'active_courses': 'Active Courses',
            'avg_grade': 'Average Grade',
            'attendance': 'Attendance',
            'announcements': 'Announcements',
            'my_timetable': 'My Timetable',
            'recent_announcements': 'Recent Announcements',
            'assigned_classes': 'Assigned Classes',
            'weekly_hours': 'Weekly Hours',
            'pending_grades': 'Pending Grades',
            'teaching_schedule': 'Teaching Schedule',
            'karabakh_uni': 'EMM',
            'smart_timetable': 'Smart Timetable',
            'footer_text': 'EMM VSA',
            'forgot_password': 'Forgot password?',
            'search_placeholder': 'Search courses by code or name...',
            'enrolled_students': 'Enrolled Students',
            'avatar': 'Avatar',
            'name': 'Name',
            'email': 'Email',
            'department': 'Department',
            'year': 'Year',
            'actions': 'Actions',
            'faculty_members': 'Faculty Members',
            'assigned_courses': 'Assigned Courses',
            'academic_courses': 'Academic Courses',
            'color': 'Color',
            'code': 'Code',
            'course_name': 'Course Name',
            'credits': 'Credits',
            'type': 'Type',
            'assigned_teacher': 'Assigned Teacher',
            'campus_rooms': 'Campus Rooms',
            'room_name': 'Room Name',
            'capacity': 'Capacity',
            'room_type': 'Room Type',
            'student_groups': 'Student Groups',
            'group_name': 'Group Name',
            'group_size': 'Group Size',
            'sys_settings': 'University Settings',
            'uni_name': 'University Name',
            'domain': 'Domain',
            'address': 'Address',
            'founded': 'Founded Year',
            'save_changes': 'Save Changes',
            'sys_details': 'System Details',
            'version': 'Version',
            'database': 'Database',
            'last_backup': 'Last Backup',
            'reset_db': 'Reset Database',
            'today': 'Today',
            'complete_timetable': 'Complete Master Timetable',
            'total_classes': 'Total Classes',
            'builder_title': 'Smart Timetable Generator',
            'builder_desc': 'Generate optimized, conflict-free academic schedules using heuristic algorithms.',
            'config_details': 'Configuration Details',
            'acad_semester': 'Academic Semester',
            'daily_start': 'Daily Start Time',
            'daily_end': 'Daily End Time',
            'max_classes': 'Max Classes per Day (Per Group)',
            'active_constraints': 'Active Constraints',
            'constraint_1': 'No teacher can teach two classes at the same time.',
            'constraint_2': 'No room can be double-booked.',
            'generated_entries': 'GENERATED ENTRIES',
            'conflicts_detected': 'CONFLICTS DETECTED',
            'current_result': 'Current Timetable Result',
            'full_view': 'Full View',
            'status': 'STATUS',
            'ok': 'OK',
            'view_details': 'View Details',
            'all_semesters': 'All Semesters',
            'lab_required': 'Lab Required',
            'lecture': 'Lecture',
            'instructor': 'Instructor',
            'enrolled': 'Enrolled',
            'available': 'Available',
            'my_course': 'My Course',
            'active_status': 'Active',
            'academic_perf': 'Academic Performance Overview',
            'grade': 'Grade',
            'pending': 'Pending',
            'passed': 'Passed',
            'failed': 'Failed',
            'in_progress': 'In Progress',
            'no_grade_data': 'No grade data available.',
            'account': 'Account',
            'personal_info': 'Personal Information',
            'full_name': 'Full Name',
            'email_addr': 'Email Address',
            'university': 'University',
            'security': 'Security',
            'security_desc': 'Manage your password and security settings.',
            'change_password': 'Change Password',
            'student_dashboard_title': 'Student Dashboard',
            'welcome_back': 'Welcome back',
            'year_prefix': 'Year',
            'current_semester': 'Current Semester',
            'no_timetable': 'No timetable generated for your group yet.'
        },
        'az': {
            'nav_dashboard': 'İdarə Paneli',
            'nav_courses': 'Fənlər',
            'nav_grades': 'Qiymətlər',
            'nav_profile': 'Profil',
            'nav_admin_dashboard': 'Admin Paneli',
            'nav_smart_timetable': 'Ağıllı Cədvəl',
            'nav_students': 'Tələbələr',
            'nav_teachers': 'Müəllimlər',
            'nav_rooms': 'Otaqlar',
            'nav_groups': 'Qruplar',
            'nav_generated_timetable': 'Yaradılmış Cədvəl',
            'nav_settings': 'Tənzimləmələr',
            'logout': 'Çıxış',
            'welcome': 'Xoş Gəldiniz',
            'sign_in_hint': 'Universitet hesabınıza daxil olun',
            'email_label': 'Universitet E-poçtu',
            'password_label': 'Şifrə',
            'sign_in_btn': 'Daxil Ol',
            'demo_acc': 'Demo Hesablar',
            'admin_dashboard_title': 'Universitet İdarəetməsi',
            'system_status': 'Sistem Vəziyyəti',
            'open_builder': 'Cədvəl Yaradıcısını Aç',
            'total_students': 'Ümumi Tələbələr',
            'total_teachers': 'Ümumi Müəllimlər',
            'conflicts': 'Münaqişələr',
            'quick_actions': 'Sürətli Əməliyyatlar',
            'generate_timetable': 'Cədvəl Yarat',
            'manage_students': 'Tələbələri İdarə Et',
            'manage_teachers': 'Müəllimləri İdarə Et',
            'manage_courses': 'Fənləri İdarə Et',
            'manage_rooms': 'Otaqları İdarə Et',
            'recent_entries': 'Son Yaradılmış Qeydlər',
            'day': 'Gün',
            'time': 'Saat',
            'course': 'Fənn',
            'teacher': 'Müəllim',
            'room': 'Otaq',
            'group': 'Qrup',
            'active_courses': 'Aktiv Fənlər',
            'avg_grade': 'Ortalama Qiymət',
            'attendance': 'Davamiyyət',
            'announcements': 'Elanlar',
            'my_timetable': 'Mənim Cədvəlim',
            'recent_announcements': 'Son Elanlar',
            'assigned_classes': 'Təyin Olunmuş Dərslər',
            'weekly_hours': 'Həftəlik Saatlar',
            'pending_grades': 'Gözləyən Qiymətlər',
            'teaching_schedule': 'Tədris Cədvəli',
            'karabakh_uni': 'EMM',
            'smart_timetable': 'Ağıllı Cədvəl',
            'footer_text': 'EMM VSA',
            'forgot_password': 'Şifrəni unutmusunuz?',
            'search_placeholder': 'Fənləri kod və ya adla axtarın...',
            'enrolled_students': 'Qeydiyyatdan Keçmiş Tələbələr',
            'avatar': 'Avatar',
            'name': 'Ad',
            'email': 'E-poçt',
            'department': 'Kafedra',
            'year': 'İl',
            'actions': 'Əməliyyatlar',
            'faculty_members': 'Fakültə Üvləri',
            'assigned_courses': 'Təyin Olunmuş Fənlər',
            'academic_courses': 'Akademik Fənlər',
            'color': 'Rəng',
            'code': 'Kod',
            'course_name': 'Fənn Adı',
            'credits': 'Kreditlər',
            'type': 'Tip',
            'assigned_teacher': 'Təyin Olunmuş Müəllim',
            'campus_rooms': 'Kampus Otaqları',
            'room_name': 'Otaq Adı',
            'capacity': 'Tutum',
            'room_type': 'Otaq Tipi',
            'student_groups': 'Tələbə Qrupları',
            'group_name': 'Qrup Adı',
            'group_size': 'Qrup Ölçüsü',
            'sys_settings': 'Universitet Tənzimləmələri',
            'uni_name': 'Universitet Adı',
            'domain': 'Domen',
            'address': 'Ünvan',
            'founded': 'Yaradılma İli',
            'save_changes': 'Dəyişiklikləri Yadda Saxla',
            'sys_details': 'Sistem Məlumatları',
            'version': 'Versiya',
            'database': 'Məlumat Bazası',
            'last_backup': 'Son Yedəkləmə',
            'reset_db': 'Məlumat Bazasını Sıfırla',
            'today': 'Bu gün',
            'complete_timetable': 'Tam Master Cədvəl',
            'total_classes': 'Ümumi Dərslər',
            'builder_title': 'Ağıllı Cədvəl Yaradıcısı',
            'builder_desc': 'Evristik alqoritmlər istifadə edərək optimallaşdırılmış, münaqişəsiz akademik cədvəllər yaradın.',
            'config_details': 'Konfiqurasiya Detalları',
            'acad_semester': 'Akademik Semestr',
            'daily_start': 'Gündəlik Başlama Saatı',
            'daily_end': 'Gündəlik Bitmə Saatı',
            'max_classes': 'Gündəlik Maksimum Dərs (Qrup üzrə)',
            'active_constraints': 'Aktiv Məhdudiyyətlər',
            'constraint_1': 'Heç bir müəllim eyni anda iki dərs deyə bilməz.',
            'constraint_2': 'Heç bir otaq eyni anda iki dəfə bron edilə bilməz.',
            'generated_entries': 'YARADILMIŞ QEYDLƏR',
            'conflicts_detected': 'AŞKARLANAN MÜNAQİŞƏLƏR',
            'current_result': 'Cari Cədvəl Nəticəsi',
            'full_view': 'Tam Görünüş',
            'status': 'STATUS',
            'ok': 'OK',
            'view_details': 'Ətraflı Bax',
            'all_semesters': 'Bütün Semestrlər',
            'lab_required': 'Laboratoriya Tələb Olunur',
            'lecture': 'Mühazirə',
            'instructor': 'Müəllim',
            'enrolled': 'Qeydiyyatdadır',
            'available': 'Açıqdır',
            'my_course': 'Mənim Dərsim',
            'active_status': 'Aktiv',
            'academic_perf': 'Akademik Performans İcmalı',
            'grade': 'Qiymət',
            'pending': 'Gözləyir',
            'passed': 'Keçdi',
            'failed': 'Kəsildi',
            'in_progress': 'Davam Edir',
            'no_grade_data': 'Heç bir qiymət məlumatı yoxdur.',
            'account': 'Hesab',
            'personal_info': 'Şəxsi Məlumatlar',
            'full_name': 'Tam Ad',
            'email_addr': 'E-poçt Ünvanı',
            'university': 'Universitet',
            'security': 'Təhlükəsizlik',
            'security_desc': 'Şifrənizi və təhlükəsizlik tənzimləmələrini idarə edin.',
            'change_password': 'Şifrəni Dəyişdir',
            'student_dashboard_title': 'Tələbə İdarə Paneli',
            'welcome_back': 'Xoş gördük',
            'year_prefix': 'İl',
            'current_semester': 'Cari Semestr',
            'no_timetable': 'Sizin qrupunuz üçün hələ cədvəl yaradılmayıb.'
        }
    };

    // Dictionary for dynamic DB words translation
    const dynamicDict = {
        'en': {
            'University Admin': 'University Admin',
            'Admin': 'Admin',
            'EMM': 'EMM',
            'Computer Science': 'Computer Science',
            'student': 'student',
            'teacher': 'teacher',
            'admin': 'admin',
            'Student': 'Student',
            'Teacher': 'Teacher',
            'Spring 2026': 'Spring 2026',
            'Fall 2025': 'Fall 2025',
            'Fall 2026': 'Fall 2026',
            'Discrete Mathematics': 'Discrete Mathematics',
            'Probability Theory and Statistics': 'Probability Theory and Statistics',
            'Computer Graphics': 'Computer Graphics',
            'Data Analysis': 'Data Analysis',
            'Web Engineering': 'Web Engineering',
            'Artificial Intelligence': 'Artificial Intelligence',
            'Database Systems': 'Database Systems',
            'TBA': 'TBA',
            'To Be Announced': 'To Be Announced',
            'Monday': 'Monday',
            'Tuesday': 'Tuesday',
            'Wednesday': 'Wednesday',
            'Thursday': 'Thursday',
            'Friday': 'Friday',
            'Saturday': 'Saturday',
            'Sunday': 'Sunday',
            'N/A': 'N/A',
            'IMPORTANT': 'IMPORTANT',
            'INFO': 'INFO',
            'WARNING': 'WARNING',
            'Midterm Exam Schedule Released': 'Midterm Exam Schedule Released',
            'Lab sessions moved to Lab-2 this week': 'Lab sessions moved to Lab-2 this week',
            'Registration deadline: May 20': 'Registration deadline: May 20'
        },
        'az': {
            'University Admin': 'Universitet Admini',
            'Admin': 'Admin',
            'EMM': 'EMM',
            'Computer Science': 'Kompüter Elmləri',
            'student': 'tələbə',
            'teacher': 'müəllim',
            'admin': 'admin',
            'Student': 'Tələbə',
            'Teacher': 'Müəllim',
            'Spring 2026': 'Yaz 2026',
            'Fall 2025': 'Payız 2025',
            'Fall 2026': 'Payız 2026',
            'Discrete Mathematics': 'Diskret Riyaziyyat',
            'Probability Theory and Statistics': 'Ehtimal Nəzəriyyəsi və Statistika',
            'Computer Graphics': 'Kompüter Qrafikası',
            'Data Analysis': 'Verilənlərin Analizi',
            'Web Engineering': 'Veb Mühəndisliyi',
            'Artificial Intelligence': 'Süni İntellekt',
            'Database Systems': 'Verilənlər Bazası Sistemləri',
            'TBA': 'Elan Ediləcək',
            'To Be Announced': 'Elan Ediləcək',
            'Monday': 'Bazar ertəsi',
            'Tuesday': 'Çərşənbə axşamı',
            'Wednesday': 'Çərşənbə',
            'Thursday': 'Cümə axşamı',
            'Friday': 'Cümə',
            'Saturday': 'Şənbə',
            'Sunday': 'Bazar',
            'N/A': 'Yoxdur',
            'IMPORTANT': 'VACİB',
            'INFO': 'MƏLUMAT',
            'WARNING': 'XƏBƏRDARLIQ',
            'Midterm Exam Schedule Released': 'Aralıq İmtahan Cədvəli Elan Edildi',
            'Lab sessions moved to Lab-2 this week': 'Bu həftə laboratoriya dərsləri Lab-2-yə keçirildi',
            'Registration deadline: May 20': 'Qeydiyyat üçün son tarix: 20 May'
        }
    };

    let currentLang = localStorage.getItem('appLang') || 'en';

    function applyTranslations(lang) {
        document.querySelectorAll('[data-i18n]').forEach(el => {
            const key = el.getAttribute('data-i18n');
            if (translations[lang] && translations[lang][key]) {
                if (el.tagName === 'INPUT' && el.hasAttribute('placeholder')) {
                    el.placeholder = translations[lang][key];
                } else {
                    const icon = el.querySelector('i');
                    if (icon) {
                        el.innerHTML = '';
                        el.appendChild(icon);
                        el.appendChild(document.createTextNode(' ' + translations[lang][key]));
                    } else {
                        el.innerText = translations[lang][key];
                    }
                }
            }
        });
        
        // Handle dynamic db words
        document.querySelectorAll('[data-i18n-dynamic]').forEach(el => {
            const originalText = el.getAttribute('data-original-text') || el.innerText.trim();
            // Store original English text if not already stored
            if (!el.getAttribute('data-original-text')) {
                el.setAttribute('data-original-text', originalText);
            }
            
            if (dynamicDict[lang] && dynamicDict[lang][originalText]) {
                const icon = el.querySelector('i');
                if (icon) {
                    el.innerHTML = '';
                    el.appendChild(icon);
                    el.appendChild(document.createTextNode(' ' + dynamicDict[lang][originalText]));
                } else {
                    el.innerText = dynamicDict[lang][originalText];
                }
            } else {
                // fallback to original
                const icon = el.querySelector('i');
                if (icon) {
                    el.innerHTML = '';
                    el.appendChild(icon);
                    el.appendChild(document.createTextNode(' ' + originalText));
                } else {
                    el.innerText = originalText;
                }
            }
        });
        
        const langToggle = document.getElementById('langToggle') || document.getElementById('langToggleLogin');
        if (langToggle) {
            langToggle.querySelector('span').innerText = lang === 'en' ? 'AZ' : 'EN';
        }
        document.documentElement.lang = lang;
    }

    // Initialize Language
    applyTranslations(currentLang);

    // Language Toggle Click
    const langToggle = document.getElementById('langToggle') || document.getElementById('langToggleLogin');
    if (langToggle) {
        langToggle.addEventListener('click', (e) => {
            e.preventDefault();
            currentLang = currentLang === 'en' ? 'az' : 'en';
            localStorage.setItem('appLang', currentLang);
            applyTranslations(currentLang);
            
            if (typeof showToast === 'function') {
                showToast(currentLang === 'az' ? 'Dil dəyişdirildi' : 'Language changed', 'success');
            }
        });
    }
});

function initializeExcelTimetableBuilder() {
    const zone = document.getElementById('upload-zone');
    if (!zone) return;

    const fileInput = document.getElementById('excel-file-input');
    const btnAnalyze = document.getElementById('btn-analyze');
    const btnGenerate = document.getElementById('btn-generate');
    const btnExport = document.getElementById('btn-export');
    const btnClearFile = document.getElementById('btn-clear-file');
    const btnClearResults = document.getElementById('btn-clear-results');
    const btnToggleView = document.getElementById('btn-toggle-view');
    const btnReuseLast = document.getElementById('btn-reuse-last');
    const generatingStatus = document.getElementById('generating-status');
    const state = { pendingFile: null, uploadedFilename: null, currentView: 'table' };

    zone.addEventListener('click', () => fileInput.click());
    zone.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            fileInput.click();
        }
    });
    zone.addEventListener('dragover', (event) => {
        event.preventDefault();
        zone.classList.add('upload-zone-active');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('upload-zone-active'));
    zone.addEventListener('drop', (event) => {
        event.preventDefault();
        zone.classList.remove('upload-zone-active');
        if (event.dataTransfer.files.length) {
            handleSelectedFile(event.dataTransfer.files[0]);
        }
    });
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length) {
            handleSelectedFile(fileInput.files[0]);
        }
    });

    if (btnClearFile) btnClearFile.addEventListener('click', clearFile);
    if (btnAnalyze) btnAnalyze.addEventListener('click', uploadAndAnalyze);
    if (btnGenerate) btnGenerate.addEventListener('click', generateTimetable);
    if (btnExport) btnExport.addEventListener('click', () => { window.location.href = '/admin/export-timetable'; });
    if (btnClearResults) btnClearResults.addEventListener('click', clearResults);
    if (btnToggleView) btnToggleView.addEventListener('click', toggleView);
    if (btnReuseLast) {
        btnReuseLast.addEventListener('click', () => {
            state.uploadedFilename = btnReuseLast.dataset.filename;
            state.pendingFile = null;
            showFilePreview(state.uploadedFilename);
            btnAnalyze.disabled = false;
            showToast('Previous upload selected for analysis.', 'success');
        });
    }

    function handleSelectedFile(file) {
        const isExcel = /\.(xlsx|xls)$/i.test(file.name);
        if (!isExcel) {
            showToast('Only .xlsx and .xls files are allowed.', 'error');
            return;
        }
        state.pendingFile = file;
        state.uploadedFilename = null;
        showFilePreview(file.name);
        btnAnalyze.disabled = false;
        document.getElementById('analysis-section').hidden = true;
        document.getElementById('results-section').hidden = true;
        renderLatestGeneratedPanel([], 'No timetable has been generated for the latest uploaded file yet.');
        btnExport.hidden = true;
    }

    function showFilePreview(filename) {
        document.getElementById('file-name-label').textContent = filename;
        document.getElementById('file-preview').hidden = false;
        const icon = zone.querySelector('.upload-zone-icon');
        const text = zone.querySelector('.upload-zone-text');
        if (icon) icon.hidden = true;
        if (text) text.hidden = true;
        zone.classList.add('has-file');
    }

    function clearFile(event) {
        if (event) event.stopPropagation();
        state.pendingFile = null;
        state.uploadedFilename = null;
        fileInput.value = '';
        document.getElementById('file-preview').hidden = true;
        const icon = zone.querySelector('.upload-zone-icon');
        const text = zone.querySelector('.upload-zone-text');
        if (icon) icon.hidden = false;
        if (text) text.hidden = false;
        zone.classList.remove('has-file');
        btnAnalyze.disabled = true;
        document.getElementById('analysis-section').hidden = true;
        clearResults();
    }

    async function uploadAndAnalyze() {
        setButtonLoading(btnAnalyze, true, 'Uploading...');
        try {
            if (state.pendingFile) {
                const formData = new FormData();
                formData.append('file', state.pendingFile);
                const uploadData = await fetchJson('/admin/upload-excel', { method: 'POST', body: formData });
                if (!uploadData.success) {
                    throw new Error(uploadData.error || 'Upload failed.');
                }
                state.uploadedFilename = uploadData.filename;
                state.pendingFile = null;
                renderLatestGeneratedPanel([], 'No timetable has been generated for the latest uploaded file yet.');
                if (uploadData.message) {
                    showToast(uploadData.message, 'success');
                }
            }

            if (!state.uploadedFilename) {
                throw new Error('Please select an Excel file first.');
            }

            setButtonLoading(btnAnalyze, true, 'Analyzing...');
            const analysisData = await fetchJson('/admin/analyze-excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: state.uploadedFilename })
            });

            if (!analysisData.success) {
                throw new Error(analysisData.error || 'Analysis failed.');
            }

            renderAnalysis(analysisData);
            renderLatestGeneratedPanel([], 'No timetable has been generated for the latest uploaded file yet.');
            showToast(analysisData.can_generate ? 'Excel analyzed successfully.' : 'Excel analyzed with issues.', analysisData.can_generate ? 'success' : 'error');
        } catch (error) {
            showToast(error.message, 'error');
        } finally {
            setButtonLoading(btnAnalyze, false, '<i class="fa-solid fa-magnifying-glass-chart"></i> Analyze Excel');
        }
    }

    function renderAnalysis(data) {
        const summary = data.summary || {};
        const cards = [
            { label: 'Teachers', value: summary.teachers || 0, icon: 'fa-chalkboard-user', color: '#0f8f5f', bg: '#ecfdf5' },
            { label: 'Rooms', value: summary.rooms || 0, icon: 'fa-door-open', color: '#2563eb', bg: '#eff6ff' },
            { label: 'Courses', value: summary.courses || 0, icon: 'fa-book-open', color: '#7c3aed', bg: '#f5f3ff' },
            { label: 'Groups', value: summary.groups || 0, icon: 'fa-users', color: '#d97706', bg: '#fffbeb' },
            { label: 'Time Slots', value: summary.time_slots || 0, icon: 'fa-clock', color: '#059669', bg: '#ecfdf5' },
            { label: 'Warnings', value: summary.warnings || 0, icon: 'fa-triangle-exclamation', color: '#b45309', bg: '#fff7ed' }
        ];

        document.getElementById('summary-cards').innerHTML = cards.map(card => `
            <div class="summary-stat-card" style="--card-bg:${card.bg};--icon-color:${card.color}">
                <div class="summary-stat-icon"><i class="fa-solid ${card.icon}"></i></div>
                <div class="summary-stat-body">
                    <div class="summary-stat-value">${escapeHtml(card.value)}</div>
                    <div class="summary-stat-label">${escapeHtml(card.label)}</div>
                </div>
            </div>
        `).join('');

        renderValidationPanel(data.errors || [], data.warnings || []);
        document.getElementById('analysis-section').hidden = false;
        btnGenerate.disabled = !data.can_generate;
        if (!data.can_generate) {
            btnGenerate.title = 'Fix validation errors before generating.';
        } else {
            btnGenerate.removeAttribute('title');
        }
    }

    function renderValidationPanel(errors, warnings) {
        const panel = document.getElementById('validation-panel');
        const header = document.getElementById('validation-header');
        const list = document.getElementById('validation-list');
        panel.hidden = false;

        if (errors.length) {
            panel.className = 'validation-panel validation-panel-error';
            header.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> ${errors.length} issue(s) found. Please fix them before generating.`;
        } else if (warnings.length) {
            panel.className = 'validation-panel validation-panel-warning';
            header.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> ${warnings.length} warning(s) found. Generation can continue.`;
        } else {
            panel.className = 'validation-panel validation-panel-success';
            header.innerHTML = '<i class="fa-solid fa-circle-check"></i> All checks passed. Ready to generate.';
        }

        const errorItems = errors.map(item => validationItem(item, 'error'));
        const warningItems = warnings.map(item => validationItem(item, 'warning'));
        list.innerHTML = [...errorItems, ...warningItems].join('');
    }

    function validationItem(message, type) {
        const icon = type === 'error' ? 'fa-circle-xmark' : 'fa-triangle-exclamation';
        return `<li class="validation-item validation-${type}">
            <i class="fa-solid ${icon}"></i>
            <span>${escapeHtml(message)}</span>
        </li>`;
    }

    async function generateTimetable() {
        if (!state.uploadedFilename) {
            showToast('Analyze an Excel file first.', 'error');
            return;
        }

        setButtonLoading(btnGenerate, true, 'Generating...');
        generatingStatus.hidden = false;
        try {
            const result = await fetchJson('/admin/generate-from-excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: state.uploadedFilename })
            });

            if (!result.success) {
                if (result.errors || result.warnings) {
                    renderValidationPanel(result.errors || [], result.warnings || []);
                }
                throw new Error(result.error || 'Generation failed.');
            }

            renderResults(result);
            showToast(`Timetable generated. ${result.summary.scheduled_sessions} sessions scheduled.`, 'success');
        } catch (error) {
            showToast(error.message, 'error');
        } finally {
            setButtonLoading(btnGenerate, false, '<i class="fa-solid fa-wand-magic-sparkles"></i> Generate Optimal Timetable');
            generatingStatus.hidden = true;
        }
    }

    function renderResults(data) {
        const summary = data.summary || {};
        const resultCards = [
            { label: 'Scheduled Sessions', value: summary.scheduled_sessions || 0, icon: 'fa-calendar-check', color: '#065f46', bg: '#ecfdf5' },
            { label: 'Unresolved Sessions', value: summary.unresolved_sessions || 0, icon: 'fa-calendar-xmark', color: '#92400e', bg: '#fef3c7' },
            { label: 'Conflicts', value: summary.conflict_count || 0, icon: 'fa-triangle-exclamation', color: '#991b1b', bg: '#fee2e2' },
            { label: 'Optimization Score', value: `${summary.optimization_score || 0}%`, icon: 'fa-gauge-high', color: '#1e40af', bg: '#eff6ff' }
        ];
        document.getElementById('result-summary-cards').innerHTML = resultCards.map(card => `
            <div class="stat-card result-stat-card" style="border-left-color:${card.color}">
                <div class="stat-icon" style="background:${card.bg};color:${card.color}"><i class="fa-solid ${card.icon}"></i></div>
                <div class="stat-details">
                    <h3>${escapeHtml(card.label)}</h3>
                    <p style="color:${card.color}">${escapeHtml(card.value)}</p>
                </div>
            </div>
        `).join('');

        const timetable = data.timetable || [];
        renderTimetableTable(timetable);
        renderWeeklyGrid(timetable);
        renderLatestGeneratedPanel(timetable);
        renderConflicts(data.conflicts || []);
        document.getElementById('results-section').hidden = false;
        btnExport.hidden = false;
        const latestPanel = document.getElementById('latest-generated-panel');
        (latestPanel || document.getElementById('results-section')).scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function getSortedTimetableEntries(entries) {
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        return [...entries].sort((a, b) => {
            const dayCompare = days.indexOf(a.day) - days.indexOf(b.day);
            return dayCompare || String(a.start_time).localeCompare(String(b.start_time));
        });
    }

    function renderTimetableTable(entries) {
        const sorted = getSortedTimetableEntries(entries);

        const tbody = document.getElementById('result-tbody');
        if (!sorted.length) {
            tbody.innerHTML = '<tr><td colspan="8" class="empty-table-message">No scheduled sessions were generated.</td></tr>';
            return;
        }

        tbody.innerHTML = sorted.map(entry => `
            <tr>
                <td><span class="badge badge-primary">${escapeHtml(entry.day)}</span></td>
                <td>${escapeHtml(entry.start_time)} - ${escapeHtml(entry.end_time)}</td>
                <td><strong>${escapeHtml(entry.course_id)}</strong> ${escapeHtml(entry.course_name)}</td>
                <td>${escapeHtml(entry.teacher_name)}</td>
                <td>${roomBadge(entry)}</td>
                <td><span class="badge badge-info">${escapeHtml(entry.group_name)}</span></td>
                <td><span class="badge ${entry.course_type === 'lab' ? 'badge-warning' : 'badge-primary'}">${escapeHtml(entry.course_type)}</span></td>
                <td><span class="badge badge-success"><i class="fa-solid fa-check"></i> ${escapeHtml(entry.status)}</span></td>
            </tr>
        `).join('');
    }

    function renderLatestGeneratedPanel(entries, emptyMessage = 'No scheduled sessions were generated.') {
        const panel = document.getElementById('latest-generated-panel');
        const tbody = document.getElementById('latest-generated-tbody');
        const count = document.getElementById('latest-generated-count');
        const actions = document.getElementById('latest-generated-actions');
        if (!panel || !tbody) return;

        const sorted = getSortedTimetableEntries(entries);
        panel.hidden = false;
        if (actions) {
            actions.hidden = sorted.length === 0;
        }
        if (count) {
            count.textContent = `${sorted.length} ${sorted.length === 1 ? 'Class' : 'Classes'}`;
        }

        if (!sorted.length) {
            tbody.innerHTML = `<tr><td colspan="8" class="empty-table-message">${escapeHtml(emptyMessage)}</td></tr>`;
            return;
        }

        tbody.innerHTML = sorted.map(entry => {
            const courseType = entry.course_type || 'lecture';
            const status = entry.status || 'scheduled';
            return `
                <tr>
                    <td><span class="badge badge-primary">${escapeHtml(entry.day)}</span></td>
                    <td>${escapeHtml(entry.start_time)} - ${escapeHtml(entry.end_time)}</td>
                    <td><strong>${escapeHtml(entry.course_id)}</strong> ${escapeHtml(entry.course_name)}</td>
                    <td>${escapeHtml(entry.teacher_name)}</td>
                    <td>${roomBadge(entry)}</td>
                    <td><span class="badge badge-info">${escapeHtml(entry.group_name)}</span></td>
                    <td><span class="badge ${courseType === 'lab' ? 'badge-warning' : 'badge-primary'}">${escapeHtml(courseType)}</span></td>
                    <td><span class="badge badge-success"><i class="fa-solid fa-check"></i> ${escapeHtml(status)}</span></td>
                </tr>
            `;
        }).join('');
    }

    function roomBadge(entry) {
        if (entry.course_type === 'lab') {
            return `<span class="badge lab-room-badge"><i class="fa-solid fa-flask"></i> ${escapeHtml(entry.room_name)}</span>`;
        }
        return escapeHtml(entry.room_name);
    }

    function renderWeeklyGrid(entries) {
        const container = document.getElementById('weekly-grid-container');
        const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        const slots = Array.from(new Set(entries.map(entry => `${entry.start_time}|${entry.end_time}`))).sort();
        const bySlotDay = {};
        entries.forEach(entry => {
            const key = `${entry.day}|${entry.start_time}|${entry.end_time}`;
            bySlotDay[key] = bySlotDay[key] || [];
            bySlotDay[key].push(entry);
        });

        if (!slots.length) {
            container.innerHTML = '<div class="weekly-grid-empty">No scheduled lessons to preview.</div>';
            container.style.gridTemplateColumns = '1fr';
            return;
        }

        container.style.gridTemplateColumns = `96px repeat(${days.length}, minmax(150px, 1fr))`;
        let html = '<div class="wg-header wg-corner">Time</div>';
        days.forEach(day => { html += `<div class="wg-header">${escapeHtml(day)}</div>`; });

        slots.forEach(slot => {
            const [start, end] = slot.split('|');
            html += `<div class="wg-time">${escapeHtml(start)}<span>${escapeHtml(end)}</span></div>`;
            days.forEach(day => {
                const items = bySlotDay[`${day}|${start}|${end}`] || [];
                html += `<div class="wg-cell">${items.map(entry => `
                    <div class="wg-class ${entry.course_type === 'lab' ? 'wg-class-lab' : ''}">
                        <div class="wg-class-code">${escapeHtml(entry.course_id)}</div>
                        <div class="wg-class-name">${escapeHtml(entry.course_name)}</div>
                        <div class="wg-class-meta"><i class="fa-solid fa-user"></i>${escapeHtml(entry.teacher_name)}</div>
                        <div class="wg-class-meta"><i class="fa-solid fa-door-open"></i>${escapeHtml(entry.room_name)} <span>${escapeHtml(entry.group_name)}</span></div>
                    </div>
                `).join('')}</div>`;
            });
        });
        container.innerHTML = html;
    }

    function renderConflicts(conflicts) {
        const panel = document.getElementById('conflict-panel');
        const list = document.getElementById('conflict-list');
        const badge = document.getElementById('conflict-count-badge');
        if (!conflicts.length) {
            panel.hidden = true;
            return;
        }
        panel.hidden = false;
        badge.textContent = `${conflicts.length} conflict(s)`;
        list.innerHTML = conflicts.map(conflict => `
            <li class="conflict-item">
                <i class="fa-solid fa-triangle-exclamation"></i>
                <span><strong>${escapeHtml(conflict.course_id)} - ${escapeHtml(conflict.course_name || '')}</strong>: ${escapeHtml(conflict.reason)}</span>
            </li>
        `).join('');
    }

    function toggleView() {
        const tableView = document.getElementById('table-view');
        const gridView = document.getElementById('grid-view');
        const icon = document.getElementById('view-toggle-icon');
        const text = document.getElementById('view-toggle-text');

        if (state.currentView === 'table') {
            tableView.hidden = true;
            gridView.hidden = false;
            icon.className = 'fa-solid fa-table';
            text.textContent = 'Table View';
            state.currentView = 'grid';
        } else {
            gridView.hidden = true;
            tableView.hidden = false;
            icon.className = 'fa-solid fa-calendar-week';
            text.textContent = 'Weekly Grid';
            state.currentView = 'table';
        }
    }

    function clearResults() {
        document.getElementById('results-section').hidden = true;
        btnExport.hidden = true;
        document.getElementById('result-summary-cards').innerHTML = '';
        document.getElementById('result-tbody').innerHTML = '';
        document.getElementById('weekly-grid-container').innerHTML = '';
        document.getElementById('conflict-panel').hidden = true;
    }
}

function initializeAdminTools() {
    const filterInputs = document.querySelectorAll('.admin-filter-input');
    filterInputs.forEach(input => {
        input.addEventListener('input', () => applyAdminFilters(input.dataset.adminFilterTarget));
    });

    document.querySelectorAll('.admin-click-row[data-detail-target]').forEach(row => {
        row.addEventListener('click', event => {
            if (event.target.closest('button, a, input, select, textarea, label')) return;
            const target = document.getElementById(row.dataset.detailTarget);
            if (!target) return;
            target.hidden = !target.hidden;
            row.classList.toggle('is-open', !target.hidden);
        });
    });

    document.querySelectorAll('[data-nested-target]').forEach(button => {
        button.addEventListener('click', event => {
            event.stopPropagation();
            const target = document.getElementById(button.dataset.nestedTarget);
            if (!target) return;
            target.hidden = !target.hidden;
        });
    });

}

function applyAdminFilters(scope) {
    if (!scope) return;
    const inputs = Array.from(document.querySelectorAll(`.admin-filter-input[data-admin-filter-target="${scope}"]`));
    const rows = document.querySelectorAll(`.admin-data-row[data-filter-scope="${scope}"]`);

    rows.forEach(row => {
        const visible = inputs.every(input => {
            const value = input.value.trim().toLowerCase();
            if (!value) return true;
            const field = input.dataset.adminFilterField;
            return String(row.dataset[field] || '').toLowerCase().includes(value);
        });
        row.hidden = !visible;
        if (!visible && row.dataset.detailTarget) {
            const detail = document.getElementById(row.dataset.detailTarget);
            if (detail) detail.hidden = true;
            row.classList.remove('is-open');
        }
    });
}

async function fetchJson(url, options) {
    const response = await fetch(url, options);
    const text = await response.text();
    let data;
    try {
        data = JSON.parse(text);
    } catch (error) {
        throw new Error('Server returned an unexpected response.');
    }
    if (!response.ok) {
        throw new Error(data.error || 'Request failed.');
    }
    return data;
}

function setButtonLoading(button, loading, content) {
    if (!button) return;
    if (loading) {
        button.dataset.previousContent = button.innerHTML;
        button.disabled = true;
        button.innerHTML = `<span class="spinner-inline"></span> ${content}`;
    } else {
        button.disabled = false;
        button.innerHTML = content || button.dataset.previousContent || button.innerHTML;
    }
}

function escapeHtml(value) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return String(value ?? '').replace(/[&<>"']/g, char => map[char]);
}

// Simple Toast Notification
function showToast(message, type = 'success') {
    const toast = document.createElement('div');
    const icon = type === 'success' ? 'fa-circle-check' : 'fa-circle-exclamation';
    toast.className = `toast-notification toast-${type === 'success' ? 'success' : 'error'}`;
    toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${escapeHtml(message)}</span>`;
    document.body.appendChild(toast);
    setTimeout(() => {
        toast.classList.add('toast-hide');
        setTimeout(() => toast.remove(), 300);
    }, 3800);
}
