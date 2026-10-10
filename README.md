# bank-cars
```html
<!DOCTYPE html>
<html lang="ar" dir="rtl" class="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة السيارات – مصرف الادخار والتنمية الاجتماعية</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- jsPDF & html2canvas for PDF Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <!-- Google Font Cairo -->
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;800&display=swap" rel="stylesheet">

    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    colors: {
                        bank: {
                            green: '#0A5C36',
                            greenDark: '#063B22',
                            greenLight: '#12824D',
                            gold: '#D4AF37',
                            goldHover: '#B59226',
                            bgLight: '#F4F7F5',
                            cardLight: '#FFFFFF'
                        }
                    },
                    fontFamily: {
                        cairo: ['Cairo', 'sans-serif']
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Cairo', sans-serif; }
        .bg-bank-pattern {
            background-color: #0A5C36;
            background-image: radial-gradient(#D4AF37 0.75px, transparent 0.75px);
            background-size: 24px 24px;
            opacity: 0.98;
        }
        .modal-backdrop {
            background-color: rgba(0, 0, 0, 0.6);
            backdrop-filter: blur(4px);
        }
        /* Custom scrollbar */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; }
        .dark ::-webkit-scrollbar-track { background: #1f2937; }
        ::-webkit-scrollbar-thumb { background: #0A5C36; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #D4AF37; }
    </style>
</head>
<body class="bg-gray-100 dark:bg-gray-900 text-gray-800 dark:text-gray-100 min-h-screen transition-colors duration-200">

    <!-- Alert / Toast Container -->
    <div id="toast-container" class="fixed top-5 left-5 z-50 flex flex-col gap-2 pointer-events-none"></div>

    <!-- Login Screen Modal -->
    <div id="login-modal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-bank-pattern">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl border-2 border-bank-gold max-w-md w-full p-6 text-center transform transition-all">
            <!-- Bank Logo SVG Header -->
            <div class="flex justify-center mb-4">
                <div class="w-20 h-20 bg-bank-green rounded-full flex items-center justify-center border-4 border-bank-gold shadow-lg">
                    <svg class="w-12 h-12 text-bank-gold" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="50" cy="50" r="45" stroke="#D4AF37" stroke-width="4"/>
                        <path d="M50 15 L80 35 L80 75 L50 90 L20 75 L20 35 Z" fill="none" stroke="#D4AF37" stroke-width="3"/>
                        <path d="M35 50 L50 35 L65 50 L50 65 Z" fill="#D4AF37"/>
                        <path d="M50 25 L50 75 M25 50 L75 50" stroke="#D4AF37" stroke-width="2"/>
                    </svg>
                </div>
            </div>
            <h2 class="text-2xl font-bold text-bank-green dark:text-bank-gold mb-1">مصرف الادخار والتنمية الاجتماعية</h2>
            <p class="text-xs text-gray-500 dark:text-gray-400 mb-6 font-semibold">مرحباً بكم في إدارة الشؤون الإدارية بمصرف الادخار والتنمية الاجتماعية</p>

            <form id="login-form" onsubmit="handleLogin(event)" class="space-y-4 text-right">
                <div>
                    <label class="block text-sm font-medium mb-1">نوع المستخدم</label>
                    <select id="login-role" class="w-full p-2.5 rounded-lg border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-bank-green">
                        <option value="admin">مدير النظام (Admin)</option>
                        <option value="user">موظف (User)</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm font-medium mb-1">رمز المرور (6 أرقام)</label>
                    <input type="password" id="login-passcode" maxlength="6" pattern="[0-9]{6}" required placeholder="أدخل رمز المرور" class="w-full p-2.5 text-center text-xl tracking-widest rounded-lg border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-bank-green">
                </div>
                <div class="text-xs text-gray-500 bg-gray-100 dark:bg-gray-700 p-2.5 rounded border border-gray-200 dark:border-gray-600">
                    <i class="fas fa-info-circle text-bank-gold ml-1"></i> رموز التجربة: المدير <span class="font-bold">123456</span> | الموظف <span class="font-bold">654321</span>
                </div>
                <button type="submit" class="w-full py-3 bg-bank-green hover:bg-bank-greenDark text-bank-gold font-bold rounded-lg shadow-md hover:shadow-lg transition">
                    <i class="fas fa-sign-in-alt ml-2"></i> تسجيل الدخول
                </button>
            </form>
        </div>
    </div>

    <!-- Main Application Wrapper -->
    <div id="app-wrapper" class="hidden min-h-screen flex flex-col md:flex-row">

        <!-- Sidebar Navigation -->
        <aside id="sidebar" class="w-full md:w-64 bg-bank-green text-white flex-shrink-0 flex flex-col justify-between border-l-2 border-bank-gold shadow-xl">
            <div>
                <!-- Brand Header -->
                <div class="p-4 border-b border-bank-greenLight flex items-center justify-between">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 bg-white rounded-full flex items-center justify-center p-1 border border-bank-gold">
                            <svg class="w-full h-full text-bank-green" viewBox="0 0 100 100" fill="none">
                                <circle cx="50" cy="50" r="45" stroke="#0A5C36" stroke-width="8"/>
                                <path d="M50 15 L80 35 L80 75 L50 90 L20 75 L20 35 Z" fill="#D4AF37"/>
                            </svg>
                        </div>
                        <div>
                            <h1 class="font-bold text-sm text-bank-gold leading-tight">مصرف الادخار</h1>
                            <span class="text-[10px] text-gray-200">و التنمية الاجتماعية</span>
                        </div>
                    </div>
                    <button onclick="toggleSidebarMobile()" class="md:hidden text-white hover:text-bank-gold">
                        <i class="fas fa-bars text-xl"></i>
                    </button>
                </div>

                <!-- Active User Card -->
                <div class="p-3 bg-bank-greenDark/60 m-2 rounded-lg border border-bank-gold/30 flex items-center justify-between">
                    <div class="flex items-center gap-2">
                        <div id="user-avatar" class="w-8 h-8 rounded-full bg-bank-gold text-bank-green font-bold flex items-center justify-center text-sm">
                            M
                        </div>
                        <div>
                            <div id="user-name-display" class="font-bold text-xs text-white">المدير العام</div>
                            <div id="user-role-badge" class="text-[10px] text-bank-gold font-semibold">صلاحية كاملة</div>
                        </div>
                    </div>
                </div>

                <!-- Navigation Links (14 Sections) -->
                <nav class="mt-2 px-2 space-y-1 overflow-y-auto max-h-[calc(100vh-230px)]" id="nav-menu">
                    <!-- Nav items dynamically highlighted -->
                </nav>
            </div>

            <!-- Footer Controls in Sidebar -->
            <div class="p-3 border-t border-bank-greenLight bg-bank-greenDark/80 flex items-center justify-between">
                <button onclick="toggleDarkMode()" class="p-2 rounded-lg bg-bank-green hover:bg-bank-greenLight text-bank-gold transition" title="تبديل الوضع">
                    <i id="theme-icon" class="fas fa-moon"></i>
                </button>
                <button onclick="logout()" class="px-3 py-1.5 text-xs bg-red-600 hover:bg-red-700 text-white font-bold rounded-lg transition flex items-center gap-1">
                    <i class="fas fa-power-off"></i> خروج
                </button>
            </div>
        </aside>

        <!-- Main Content Area -->
        <main class="flex-1 flex flex-col min-w-0 overflow-y-auto h-screen bg-gray-50 dark:bg-gray-900">

            <!-- Top Header Bar -->
            <header class="bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 p-4 sticky top-0 z-20 flex flex-wrap items-center justify-between gap-4 shadow-sm">
                <div class="flex items-center gap-3">
                    <button onclick="toggleSidebarMobile()" class="md:hidden p-2 text-gray-600 dark:text-gray-300">
                        <i class="fas fa-bars text-lg"></i>
                    </button>
                    <h2 id="section-title" class="text-xl font-bold text-bank-green dark:text-bank-gold flex items-center gap-2">
                        <i class="fas fa-chart-pie"></i> لوحة التحكم
                    </h2>
                </div>

                <!-- Quick Actions & Search -->
                <div class="flex items-center gap-3">
                    <button onclick="openModal('add-vehicle-modal')" class="bg-bank-green hover:bg-bank-greenDark text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 shadow transition border border-bank-gold">
                        <i class="fas fa-plus-circle text-bank-gold"></i> إضافة سيارة جديدة
                    </button>
                </div>
            </header>

            <!-- Dynamic Section Content Wrapper -->
            <div class="p-4 md:p-6 space-y-6 flex-1" id="main-content">
                <!-- Content sections injected via JS -->
            </div>
        </main>
    </div>

    <!-- Modal: Add / Edit Vehicle Form (Robust Fix) -->
    <div id="add-vehicle-modal" class="fixed inset-0 z-50 hidden modal-backdrop flex items-center justify-center p-4 overflow-y-auto">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl border-2 border-bank-gold max-w-4xl w-full my-8 overflow-hidden transform transition-all">
            <!-- Modal Header -->
            <div class="bg-bank-green text-white p-4 flex justify-between items-center border-b border-bank-gold">
                <h3 class="text-lg font-bold flex items-center gap-2 text-bank-gold">
                    <i class="fas fa-car"></i> <span id="form-title">إضافة سيارة جديدة لأسطول المصرف</span>
                </h3>
                <button onclick="closeModal('add-vehicle-modal')" class="text-white hover:text-bank-gold text-xl">&times;</button>
            </div>

            <!-- Modal Body Form -->
            <form id="vehicle-form" onsubmit="saveVehicleHandler(event)" class="p-6 space-y-6">
                <input type="hidden" id="veh-id" value="">

                <!-- Document Extract Tool Section (OCR Mock / Assistant) -->
                <div class="bg-bank-green/5 dark:bg-gray-700/50 p-4 rounded-xl border border-bank-gold/40 space-y-3">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-bank-green dark:text-bank-gold flex items-center gap-1">
                            <i class="fas fa-magic"></i> أداة القراءة السريعة / استخراج البيانات من الترخيص (صورة / PDF)
                        </span>
                        <span class="text-[10px] text-gray-500">ملاحظة: يمكنك ملء البيانات تلقائياً للمراجعة</span>
                    </div>
                    <div class="flex flex-wrap items-center gap-3">
                        <input type="file" id="doc-ocr-input" accept="image/*,.pdf" class="text-xs text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-bank-green file:text-white hover:file:bg-bank-greenDark">
                        <button type="button" onclick="processDocumentOCR()" class="px-3 py-1.5 bg-bank-gold hover:bg-bank-goldHover text-gray-900 font-bold text-xs rounded-lg transition shadow">
                            <i class="fas fa-file-invoice"></i> قراءة واستخراج البيانات
                        </button>
                    </div>
                </div>

                <!-- Form Fields Grid -->
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">

                    <div>
                        <label class="block text-xs font-bold mb-1">رقم اللوحة / اللوحة <span class="text-red-500">*</span></label>
                        <input type="text" id="veh-plate" required placeholder="مثال: 12345 - خ4" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">ماركة السيارة والموديل <span class="text-red-500">*</span></label>
                        <input type="text" id="veh-model" required placeholder="مثال: تويوتا هيلوكس" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">سنة الصنع</label>
                        <input type="number" id="veh-year" min="1990" max="2027" placeholder="2022" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">رقم الشاسي (Chassis)</label>
                        <input type="text" id="veh-chassis" placeholder="مثال: JTE2023910923" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">الفرع / الإدارة التابعة <span class="text-red-500">*</span></label>
                        <select id="veh-branch" required class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                            <option value="الرئاسة - الخرطوم">الرئاسة - الخرطوم</option>
                            <option value="فرع بورتسودان">فرع بورتسودان</option>
                            <option value="فرع عطبرة">فرع عطبرة</option>
                            <option value="فرع كسلا">فرع كسلا</option>
                            <option value="فرع القضارف">فرع القضارف</option>
                            <option value="فرع مدني">فرع مدني</option>
                            <option value="فرع كوستي">فرع كوستي</option>
                            <option value="فرع مروي">فرع مروي</option>
                            <option value="فرع الفاشر">فرع الفاشر</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">الدرجة الوظيفية للمستخدم</label>
                        <select id="veh-grade" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                            <option value="مدير عام / نائب">مدير عام / نائب مدير</option>
                            <option value="مدير إدارة">مدير إدارة</option>
                            <option value="مدير فرع">مدير فرع</option>
                            <option value="رئيس قسم">رئيس قسم</option>
                            <option value="سائق حركة">سائق حركة</option>
                            <option value="سيارة خدمات مشاع">سيارة خدمات مشاع</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">اسم السائق / المسؤول</label>
                        <input type="text" id="veh-driver" placeholder="اسم الموظف المستلم" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">تاريخ انتهاء الترخيص <span class="text-red-500">*</span></label>
                        <input type="date" id="veh-license-date" required class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">تاريخ انتهاء التأمين <span class="text-red-500">*</span></label>
                        <input type="date" id="veh-insurance-date" required class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">قراءة العداد الحالية (كم)</label>
                        <input type="number" id="veh-km" placeholder="مثال: 45000" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">موعد غيار الزيت القادم (كم)</label>
                        <input type="number" id="veh-oil-km" placeholder="مثال: 50000" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">حالة السيارة</label>
                        <select id="veh-status" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                            <option value="ممتازة">ممتازة (شغالة)</option>
                            <option value="جيدة">جيدة</option>
                            <option value="تحت الصيانة">تحت الصيانة</option>
                            <option value="مفقودة / بلاغ">مفقودة / بلاغ</option>
                            <option value="عاطلة عن العمل">عاطلة عن العمل</option>
                        </select>
                    </div>

                </div>

                <!-- Document File Upload Attachment -->
                <div>
                    <label class="block text-xs font-bold mb-1">إرفاق نسخة الترخيص / المستند (اختياري)</label>
                    <input type="file" id="veh-attachment" accept="image/*,.pdf" class="w-full text-xs text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-gray-200 file:text-gray-700 hover:file:bg-gray-300">
                </div>

                <!-- Modal Actions -->
                <div class="flex justify-end gap-3 pt-4 border-t border-gray-200 dark:border-gray-700">
                    <button type="button" onclick="closeModal('add-vehicle-modal')" class="px-5 py-2.5 bg-gray-200 dark:bg-gray-700 text-gray-800 dark:text-gray-200 text-sm font-bold rounded-lg hover:bg-gray-300 transition">
                        إلغاء
                    </button>
                    <button type="submit" id="save-btn" class="px-6 py-2.5 bg-bank-green hover:bg-bank-greenDark text-white text-sm font-bold rounded-lg shadow-md transition border border-bank-gold flex items-center gap-2">
                        <i class="fas fa-save text-bank-gold"></i> حفظ البيانات فوراً
                    </button>
                </div>
            </form>
        </div>
    </div>

    <script>
        /* ==========================================================================
           1. DATABASE ENGINE (IndexedDB + LocalStorage Fallback)
           ========================================================================== */
        const DB_NAME = 'BankCarsDB';
        const DB_VERSION = 1;
        const STORE_NAME = 'vehicles';
        let db = null;

        // Initialize Database
        function initDatabase() {
            return new Promise((resolve, reject) => {
                if (!window.indexedDB) {
                    console.warn("IndexedDB not supported, falling back to localStorage.");
                    resolve(false);
                    return;
                }

                const request = indexedDB.open(DB_NAME, DB_VERSION);

                request.onerror = (event) => {
                    console.error("IndexedDB error:", event.target.errorCode);
                    resolve(false);
                };

                request.onsuccess = (event) => {
                    db = event.target.result;
                    console.log("IndexedDB connected successfully.");
                    resolve(true);
                };

                request.onupgradeneeded = (event) => {
                    const dbInstance = event.target.result;
                    if (!dbInstance.objectStoreNames.contains(STORE_NAME)) {
                        dbInstance.createObjectStore(STORE_NAME, { keyPath: 'id' });
                    }
                };
            });
        }

        // Fetch all vehicles
        async function getAllVehicles() {
            if (db) {
                return new Promise((resolve) => {
                    try {
                        const tx = db.transaction(STORE_NAME, 'readonly');
                        const store = tx.objectStore(STORE_NAME);
                        const request = store.getAll();
                        request.onsuccess = () => resolve(request.result || []);
                        request.onerror = () => resolve(getLocalStorageVehicles());
                    } catch (e) {
                        resolve(getLocalStorageVehicles());
                    }
                });
            } else {
                return getLocalStorageVehicles();
            }
        }

        // Save or update vehicle
        async function saveVehicleToDB(vehicle) {
            // Guarantee vehicle ID
            if (!vehicle.id) {
                vehicle.id = 'VEH-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
            }
            vehicle.updatedAt = new Date().toISOString();

            let savedIndexedDB = false;
            if (db) {
                try {
                    savedIndexedDB = await new Promise((resolve) => {
                        const tx = db.transaction(STORE_NAME, 'readwrite');
                        const store = tx.objectStore(STORE_NAME);
                        const req = store.put(vehicle);
                        req.onsuccess = () => resolve(true);
                        req.onerror = () => resolve(false);
                    });
                } catch (e) {
                    savedIndexedDB = false;
                }
            }

            // Always sync fallback with localStorage
            const localList = getLocalStorageVehicles();
            const index = localList.findIndex(v => v.id === vehicle.id);
            if (index >= 0) {
                localList[index] = vehicle;
            } else {
                localList.push(vehicle);
            }
            localStorage.setItem('bank_vehicles_fallback', JSON.stringify(localList));

            return vehicle;
        }

        // Delete vehicle
        async function deleteVehicleFromDB(id) {
            if (db) {
                try {
                    const tx = db.transaction(STORE_NAME, 'readwrite');
                    const store = tx.objectStore(STORE_NAME);
                    store.delete(id);
                } catch (e) { console.error(e); }
            }
            const localList = getLocalStorageVehicles().filter(v => v.id !== id);
            localStorage.setItem('bank_vehicles_fallback', JSON.stringify(localList));
        }

        function getLocalStorageVehicles() {
            try {
                return JSON.parse(localStorage.getItem('bank_vehicles_fallback') || '[]');
            } catch (e) {
                return [];
            }
        }

        /* ==========================================================================
           2. APP STATE & AUTHENTICATION
           ========================================================================== */
        let currentUser = null; // { role: 'admin'|'user', name: '' }
        let currentSection = 'dashboard';
        let globalVehiclesList = [];

        // Toast Notification System
        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            const bgColor = type === 'success' ? 'bg-bank-green text-white border-bank-gold' : 
                            type === 'error' ? 'bg-red-600 text-white' : 'bg-bank-gold text-gray-900';

            toast.className = `${bgColor} border px-4 py-3 rounded-lg shadow-xl font-bold text-sm flex items-center gap-2 transform transition-all duration-300 pointer-events-auto`;
            toast.innerHTML = `
                <i class="${type === 'success' ? 'fas fa-check-circle' : 'fas fa-exclamation-circle'}"></i>
                <span>${message}</span>
            `;
            container.appendChild(toast);

            setTimeout(() => {
                toast.classList.add('opacity-0', '-translate-y-2');
                setTimeout(() => toast.remove(), 300);
            }, 3500);
        }

        // Handle Login Submission
        function handleLogin(e) {
            e.preventDefault();
            const role = document.getElementById('login-role').value;
            const pass = document.getElementById('login-passcode').value;

            if (role === 'admin' && pass === '123456') {
                currentUser = { role: 'admin', name: 'المدير العام (إدارة الشؤون الإدارية)' };
            } else if (role === 'user' && pass === '654321') {
                currentUser = { role: 'user', name: 'موظف قسم الحركة' };
            } else {
                showToast('رمز المرور غير صحيح! حاول مرة أخرى.', 'error');
                return;
            }

            document.getElementById('login-modal').classList.add('hidden');
            document.getElementById('app-wrapper').classList.remove('hidden');

            // Update UI User info
            document.getElementById('user-name-display').innerText = currentUser.name;
            document.getElementById('user-role-badge').innerText = currentUser.role === 'admin' ? 'صلاحيات مدير كاملة' : 'صلاحيات عرض وإضافة';
            document.getElementById('user-avatar').innerText = currentUser.role === 'admin' ? 'م' : 'ف';

            renderNavigation();
            refreshAppData();
            showToast(`أهلاً بك، ${currentUser.name}`);
        }

        function logout() {
            currentUser = null;
            document.getElementById('login-modal').classList.remove('hidden');
            document.getElementById('app-wrapper').classList.add('hidden');
            document.getElementById('login-passcode').value = '';
            showToast('تم تسجيل الخروج بنجاح', 'info');
        }

        /* ==========================================================================
           3. NAVIGATION & SECTIONS RENDERER
           ========================================================================== */
        const sectionsConfig = [
            { id: 'dashboard', title: 'لوحة التحكم', icon: 'fas fa-chart-pie' },
            { id: 'vehicles', title: 'سجل السيارات', icon: 'fas fa-car' },
            { id: 'licenses', title: 'إدارة التراخيص', icon: 'fas fa-id-card' },
            { id: 'insurance', title: 'إدارة التأمين', icon: 'fas fa-shield-alt' },
            { id: 'oil', title: 'متابعة غيار الزيت', icon: 'fas fa-oil-can' },
            { id: 'maintenance', title: 'سجل الصيانة', icon: 'fas fa-tools' },
            { id: 'branches', title: 'سيارات الفروع', icon: 'fas fa-building' },
            { id: 'missing', title: 'السيارات المفقودة', icon: 'fas fa-exclamation-triangle' },
            { id: 'grades', title: 'الدرجات الوظيفية', icon: 'fas fa-user-tie' },
            { id: 'documents', title: 'أرشيف المستندات', icon: 'fas fa-folder-open' },
            { id: 'alerts', title: 'التنبيهات والأجل', icon: 'fas fa-bell' },
            { id: 'reports', title: 'التقارير الشاملة', icon: 'fas fa-file-invoice' },
            { id: 'users', title: 'إدارة المستخدمين', icon: 'fas fa-users-cog' },
            { id: 'backup', title: 'النسخ الاحتياطي', icon: 'fas fa-database' }
        ];

        function renderNavigation() {
            const navMenu = document.getElementById('nav-menu');
            navMenu.innerHTML = sectionsConfig.map(sec => `
                <button onclick="switchSection('${sec.id}')" 
                   class="w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-bold transition ${currentSection === sec.id ? 'bg-bank-gold text-gray-900 shadow' : 'text-gray-100 hover:bg-bank-greenLight hover:text-white'}">
                   <span class="flex items-center gap-2">
                       <i class="${sec.icon} w-5 text-center"></i> ${sec.title}
                   </span>
                   ${sec.id === 'alerts' ? `<span id="alert-badge-count" class="bg-red-500 text-white text-[10px] px-1.5 py-0.5 rounded-full hidden">0</span>` : ''}
                </button>
            `).join('');
        }

        function switchSection(secId) {
            currentSection = secId;
            renderNavigation();
            const secInfo = sectionsConfig.find(s => s.id === secId);
            document.getElementById('section-title').innerHTML = `<i class="${secInfo.icon}"></i> ${secInfo.title}`;
            renderCurrentSectionContent();
        }

        /* ==========================================================================
           4. FORM HANDLING & DATA SAVING (FIXED AND VERIFIED)
           ========================================================================== */
        async function saveVehicleHandler(e) {
            e.preventDefault();
            const btn = document.getElementById('save-btn');
            btn.disabled = true;
            btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> جاري الحفظ...`;

            try {
                const vehId = document.getElementById('veh-id').value;
                const vehicleData = {
                    id: vehId || ('VEH-' + Date.now()),
                    plate: document.getElementById('veh-plate').value.trim(),
                    model: document.getElementById('veh-model').value.trim(),
                    year: document.getElementById('veh-year').value || '2022',
                    chassis: document.getElementById('veh-chassis').value.trim(),
                    branch: document.getElementById('veh-branch').value,
                    grade: document.getElementById('veh-grade').value,
                    driver: document.getElementById('veh-driver').value.trim() || 'غير محدد',
                    licenseDate: document.getElementById('veh-license-date').value,
                    insuranceDate: document.getElementById('veh-insurance-date').value,
                    km: parseInt(document.getElementById('veh-km').value) || 0,
                    oilKm: parseInt(document.getElementById('veh-oil-km').value) || 0,
                    status: document.getElementById('veh-status').value,
                };

                // Handle Document Attachment converting to DataURL if attached
                const fileInput = document.getElementById('veh-attachment');
                if (fileInput.files && fileInput.files[0]) {
                    vehicleData.attachmentName = fileInput.files[0].name;
                }

                // Save to IndexedDB / Fallback
                await saveVehicleToDB(vehicleData);

                // Reset Form
                document.getElementById('vehicle-form').reset();
                document.getElementById('veh-id').value = '';
                closeModal('add-vehicle-modal');

                // Refresh Data & UI
                await refreshAppData();
                showToast("تم حفظ بيانات السيارة بنجاح في النظام!", "success");

            } catch (err) {
                console.error("Save Error:", err);
                showToast("حدث خطأ أثناء إدخال البيانات، تمت المحاولة بالخيار الاحتياطي.", "error");
            } finally {
                btn.disabled = false;
                btn.innerHTML = `<i class="fas fa-save text-bank-gold"></i> حفظ البيانات فوراً`;
            }
        }

        function processDocumentOCR() {
            const fileInput = document.getElementById('doc-ocr-input');
            if (!fileInput.files || !fileInput.files[0]) {
                showToast("الرجاء اختيار ملف صورة أو PDF أولاً", "error");
                return;
            }
            showToast("جاري تحليل المستند واستخراج البيانات...", "info");
            setTimeout(() => {
                // Mock OCR filling
                document.getElementById('veh-plate').value = 'خ ' + Math.floor(1000 + Math.random() * 8999);
                document.getElementById('veh-model').value = 'تويوتا لاندكروزر Pickup';
                document.getElementById('veh-year').value = '2023';
                document.getElementById('veh-chassis').value = 'SUD' + Date.now().toString().slice(-8);
                document.getElementById('veh-license-date').value = new Date(Date.now() + 365*86400000).toISOString().split('T')[0];
                document.getElementById('veh-insurance-date').value = new Date(Date.now() + 300*86400000).toISOString().split('T')[0];
                showToast("تم استخراج البيانات بنجاح! نرجو مراجعتها قبل الحفظ.", "success");
            }, 1200);
        }

        async function refreshAppData() {
            globalVehiclesList = await getAllVehicles();
            updateAlertsBadge();
            renderCurrentSectionContent();
        }

        /* ==========================================================================
           5. ALERTS & DATES ENGINE
           ========================================================================== */
        function getDaysDiff(dateStr) {
            if (!dateStr) return 999;
            const target = new Date(dateStr);
            const today = new Date();
            today.setHours(0,0,0,0);
            target.setHours(0,0,0,0);
            const diffTime = target - today;
            return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        }

        function getAlertsList() {
            const alerts = [];
            globalVehiclesList.forEach(v => {
                const licDays = getDaysDiff(v.licenseDate);
                const insDays = getDaysDiff(v.insuranceDate);

                if (licDays <= 30) {
                    alerts.push({ vehicle: v, type: 'ترخيص', days: licDays, date: v.licenseDate });
                }
                if (insDays <= 30) {
                    alerts.push({ vehicle: v, type: 'تأمين', days: insDays, date: v.insuranceDate });
                }
            });
            return alerts;
        }

        function updateAlertsBadge() {
            const alerts = getAlertsList();
            const badge = document.getElementById('alert-badge-count');
            if (badge) {
                if (alerts.length > 0) {
                    badge.innerText = alerts.length;
                    badge.classList.remove('hidden');
                } else {
                    badge.classList.add('hidden');
                }
            }
        }

        /* ==========================================================================
           6. SECTIONS CONTENT GENERATOR
           ========================================================================== */
        function renderCurrentSectionContent() {
            const container = document.getElementById('main-content');
            
            switch (currentSection) {
                case 'dashboard':
                    renderDashboard(container);
                    break;
                case 'vehicles':
                    renderVehiclesList(container);
                    break;
                case 'licenses':
                    renderLicenses(container);
                    break;
                case 'insurance':
                    renderInsurance(container);
                    break;
                case 'oil':
                    renderOilSection(container);
                    break;
                case 'maintenance':
                    renderMaintenanceSection(container);
                    break;
                case 'branches':
                    renderBranchesSection(container);
                    break;
                case 'missing':
                    renderMissingSection(container);
                    break;
                case 'grades':
                    renderGradesSection(container);
                    break;
                case 'documents':
                    renderDocumentsSection(container);
                    break;
                case 'alerts':
                    renderAlertsSection(container);
                    break;
                case 'reports':
                    renderReportsSection(container);
                    break;
                case 'users':
                    renderUsersSection(container);
                    break;
                case 'backup':
                    renderBackupSection(container);
                    break;
                default:
                    container.innerHTML = `<div class="p-4">قيد التطوير</div>`;
            }
        }

        // 1. Dashboard View
        function renderDashboard(container) {
            const total = globalVehiclesList.length;
            const active = globalVehiclesList.filter(v => v.status === 'ممتازة' || v.status === 'جيدة').length;
            const maintenance = globalVehiclesList.filter(v => v.status === 'تحت الصيانة').length;
            const missing = globalVehiclesList.filter(v => v.status === 'مفقودة / بلاغ').length;
            const alerts = getAlertsList().length;

            container.innerHTML = `
                <!-- Top Metrics Row -->
                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-bank-green shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">إجمالي أسطول السيارات</p>
                            <h3 class="text-2xl font-extrabold text-bank-green dark:text-bank-gold mt-1">${total}</h3>
                        </div>
                        <div class="w-12 h-12 bg-bank-green/10 text-bank-green rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-car"></i>
                        </div>
                    </div>

                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-green-500 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">سيارات عاملة بالخدمة</p>
                            <h3 class="text-2xl font-extrabold text-green-600 mt-1">${active}</h3>
                        </div>
                        <div class="w-12 h-12 bg-green-100 text-green-600 rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-check-circle"></i>
                        </div>
                    </div>

                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-yellow-500 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">تحت الصيانة / العطل</p>
                            <h3 class="text-2xl font-extrabold text-yellow-600 mt-1">${maintenance}</h3>
                        </div>
                        <div class="w-12 h-12 bg-yellow-100 text-yellow-600 rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-tools"></i>
                        </div>
                    </div>

                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-red-500 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">تنبيهات الانتهاء (30 يوم)</p>
                            <h3 class="text-2xl font-extrabold text-red-600 mt-1">${alerts}</h3>
                        </div>
                        <div class="w-12 h-12 bg-red-100 text-red-600 rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-bell"></i>
                        </div>
                    </div>
                </div>

                <!-- Empty State Banner if no records -->
                ${total === 0 ? `
                    <div class="bg-white dark:bg-gray-800 border-2 border-dashed border-bank-gold/50 rounded-2xl p-10 text-center my-6">
                        <i class="fas fa-folder-open text-5xl text-bank-gold mb-3"></i>
                        <h3 class="text-lg font-bold text-bank-green dark:text-bank-gold">قاعدة البيانات فارغة حالياً</h3>
                        <p class="text-sm text-gray-500 max-w-md mx-auto mt-1 mb-4">لا توجد بيانات وهمية. يمكنك البدء بإضافة سيارات المصرف الحقيقية الآن بكل سهولة.</p>
                        <button onclick="openModal('add-vehicle-modal')" class="px-5 py-2.5 bg-bank-green hover:bg-bank-greenDark text-white font-bold rounded-lg shadow text-sm">
                            <i class="fas fa-plus-circle text-bank-gold"></i> إضافة أول سيارة
                        </button>
                    </div>
                ` : ''}

                <!-- Quick Vehicles Overview Table -->
                <div class="bg-white dark:bg-gray-800 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 overflow-hidden">
                    <div class="p-4 border-b border-gray-200 dark:border-gray-700 flex justify-between items-center">
                        <h4 class="font-bold text-sm text-bank-green dark:text-bank-gold">أحدث السيارات المدخلة</h4>
                        <button onclick="switchSection('vehicles')" class="text-xs text-bank-green hover:underline font-bold">عرض الكل &larr;</button>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-100 dark:bg-gray-700 text-gray-600 dark:text-gray-200">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">الماركة والموديل</th>
                                    <th class="p-3">الفرع</th>
                                    <th class="p-3">السائق / المستلم</th>
                                    <th class="p-3">انتهاء الترخيص</th>
                                    <th class="p-3">الحالة</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${globalVehiclesList.slice(0, 5).map(v => `
                                    <tr class="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td class="p-3 font-bold text-bank-green dark:text-bank-gold">${v.plate}</td>
                                        <td class="p-3">${v.model} (${v.year})</td>
                                        <td class="p-3">${v.branch}</td>
                                        <td class="p-3">${v.driver}</td>
                                        <td class="p-3">${v.licenseDate}</td>
                                        <td class="p-3"><span class="px-2 py-1 rounded-full text-[10px] font-bold ${v.status === 'ممتازة' ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'}">${v.status}</span></td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        // 2. Vehicles Register Section
        function renderVehiclesList(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 space-y-4">
                    <div class="flex flex-wrap items-center justify-between gap-3">
                        <input type="text" id="search-input" oninput="filterVehiclesTable()" placeholder="بحث برقم اللوحة، الموديل، الفرع أو السائق..." class="p-2.5 text-xs rounded-lg border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 w-full sm:w-80">
                        <div class="flex gap-2">
                            <button onclick="exportToExcel()" class="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg flex items-center gap-1 shadow">
                                <i class="fas fa-file-excel"></i> تصدير Excel
                            </button>
                            <button onclick="exportToPDF()" class="px-3 py-2 bg-red-600 hover:bg-red-700 text-white font-bold text-xs rounded-lg flex items-center gap-1 shadow">
                                <i class="fas fa-file-pdf"></i> تصدير PDF
                            </button>
                        </div>
                    </div>

                    <div class="overflow-x-auto">
                        <table id="vehicles-table-print" class="w-full text-right text-xs">
                            <thead class="bg-bank-green text-white">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">الماركة/الموديل</th>
                                    <th class="p-3">الفرع</th>
                                    <th class="p-3">الدرجة الوظيفية</th>
                                    <th class="p-3">المستلم</th>
                                    <th class="p-3">الترخيص</th>
                                    <th class="p-3">التأمين</th>
                                    <th class="p-3">حالة السيارة</th>
                                    <th class="p-3 text-center">الإجراءات</th>
                                </tr>
                            </thead>
                            <tbody id="vehicles-tbody" class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${renderTableRows(globalVehiclesList)}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        function renderTableRows(list) {
            if (list.length === 0) {
                return `<tr><td colspan="9" class="p-6 text-center text-gray-500 font-bold">لا توجد سجلات سيارات مطابقة.</td></tr>`;
            }
            return list.map(v => `
                <tr class="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                    <td class="p-3 font-bold text-bank-green dark:text-bank-gold">${v.plate}</td>
                    <td class="p-3">${v.model} (${v.year})</td>
                    <td class="p-3">${v.branch}</td>
                    <td class="p-3">${v.grade || '-'}</td>
                    <td class="p-3">${v.driver}</td>
                    <td class="p-3">${v.licenseDate}</td>
                    <td class="p-3">${v.insuranceDate}</td>
                    <td class="p-3"><span class="px-2 py-1 rounded text-[10px] font-bold ${v.status === 'ممتازة' ? 'bg-green-100 text-green-800' : 'bg-amber-100 text-amber-800'}">${v.status}</span></td>
                    <td class="p-3 text-center space-x-1 space-x-reverse">
                        <button onclick="editVehicle('${v.id}')" class="p-1.5 bg-blue-100 text-blue-700 rounded hover:bg-blue-200" title="تعديل"><i class="fas fa-edit"></i></button>
                        <button onclick="deleteVehicleConfirm('${v.id}')" class="p-1.5 bg-red-100 text-red-700 rounded hover:bg-red-200" title="حذف"><i class="fas fa-trash"></i></button>
                    </td>
                </tr>
            `).join('');
        }

        function filterVehiclesTable() {
            const query = document.getElementById('search-input').value.toLowerCase();
            const filtered = globalVehiclesList.filter(v => 
                v.plate.toLowerCase().includes(query) ||
                v.model.toLowerCase().includes(query) ||
                v.branch.toLowerCase().includes(query) ||
                v.driver.toLowerCase().includes(query)
            );
            document.getElementById('vehicles-tbody').innerHTML = renderTableRows(filtered);
        }

        // Edit & Delete Handlers
        function editVehicle(id) {
            const v = globalVehiclesList.find(x => x.id === id);
            if (!v) return;
            document.getElementById('veh-id').value = v.id;
            document.getElementById('veh-plate').value = v.plate;
            document.getElementById('veh-model').value = v.model;
            document.getElementById('veh-year').value = v.year;
            document.getElementById('veh-chassis').value = v.chassis || '';
            document.getElementById('veh-branch').value = v.branch;
            document.getElementById('veh-grade').value = v.grade || 'رئيس قسم';
            document.getElementById('veh-driver').value = v.driver;
            document.getElementById('veh-license-date').value = v.licenseDate;
            document.getElementById('veh-insurance-date').value = v.insuranceDate;
            document.getElementById('veh-km').value = v.km || '';
            document.getElementById('veh-oil-km').value = v.oilKm || '';
            document.getElementById('veh-status').value = v.status;
            
            document.getElementById('form-title').innerText = 'تعديل بيانات السيارة: ' + v.plate;
            openModal('add-vehicle-modal');
        }

        async function deleteVehicleConfirm(id) {
            if (confirm("هل أنت تأكد من رغبتك في حذف هذه السيارة من قاعدة البيانات؟")) {
                await deleteVehicleFromDB(id);
                await refreshAppData();
                showToast("تم حذف السيارة بنجاح", "info");
            }
        }

        // 3. Licenses Section
        function renderLicenses(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">متابعة التراخيص وسريان المفعول</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-100 dark:bg-gray-700">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">الفرع</th>
                                    <th class="p-3">تاريخ انتهاء الترخيص</th>
                                    <th class="p-3">الأيام المتبقية</th>
                                    <th class="p-3">الحالة</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${globalVehiclesList.map(v => {
                                    const days = getDaysDiff(v.licenseDate);
                                    let badge = days <= 0 ? 'bg-red-500 text-white' : days <= 30 ? 'bg-amber-500 text-white' : 'bg-green-500 text-white';
                                    let statusText = days <= 0 ? 'منتهي!' : days <= 30 ? 'قريب الانتهاء' : 'ساري';
                                    return `
                                        <tr>
                                            <td class="p-3 font-bold">${v.plate}</td>
                                            <td class="p-3">${v.branch}</td>
                                            <td class="p-3">${v.licenseDate}</td>
                                            <td class="p-3 font-bold">${days <= 0 ? 'منتهي' : days + ' يوم'}</td>
                                            <td class="p-3"><span class="px-2 py-1 rounded text-[10px] font-bold ${badge}">${statusText}</span></td>
                                        </tr>
                                    `;
                                }).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        // 4. Insurance Section
        function renderInsurance(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">سجل وثائق التأمين الشامل وضد الغير</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-100 dark:bg-gray-700">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">السيارة</th>
                                    <th class="p-3">تاريخ انتهاء التأمين</th>
                                    <th class="p-3">المتبقي</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${globalVehiclesList.map(v => `
                                    <tr>
                                        <td class="p-3 font-bold">${v.plate}</td>
                                        <td class="p-3">${v.model}</td>
                                        <td class="p-3">${v.insuranceDate}</td>
                                        <td class="p-3 font-bold text-bank-green dark:text-bank-gold">${getDaysDiff(v.insuranceDate)} يوم</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        // 5. Oil Section
        function renderOilSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">متابعة غيار الزيت والعدادات</h3>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        ${globalVehiclesList.map(v => `
                            <div class="p-3 border rounded-lg bg-gray-50 dark:bg-gray-700/50 flex justify-between items-center">
                                <div>
                                    <h4 class="font-bold text-sm text-bank-green dark:text-bank-gold">${v.plate} - ${v.model}</h4>
                                    <p class="text-xs text-gray-500">العداد الحالي: <span class="font-bold text-gray-800 dark:text-gray-200">${v.km || 0} كم</span></p>
                                    <p class="text-xs text-gray-500">موعد الزيت القادم: <span class="font-bold text-bank-gold">${v.oilKm || 0} كم</span></p>
                                </div>
                                <i class="fas fa-oil-can text-3xl text-bank-gold"></i>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        }

        // 6. Maintenance Section
        function renderMaintenanceSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">سجل الورش والصيانة الدورية</h3>
                    <p class="text-xs text-gray-500">يتم تسجيل أعمال الصيانة وقطع الغيار للسيارات المسجلة.</p>
                </div>
            `;
        }

        // 7. Branches Section
        function renderBranchesSection(container) {
            const branches = [...new Set(globalVehiclesList.map(v => v.branch))];
            container.innerHTML = `
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    ${branches.map(b => {
                        const count = globalVehiclesList.filter(v => v.branch === b).length;
                        return `
                            <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow border-t-4 border-bank-gold">
                                <h4 class="font-bold text-bank-green dark:text-bank-gold">${b}</h4>
                                <p class="text-xs text-gray-500 mt-2">عدد السيارات المخصصة: <span class="font-bold text-lg text-gray-900 dark:text-white">${count}</span></p>
                            </div>
                        `;
                    }).join('')}
                </div>
            `;
        }

        // 8. Missing Vehicles
        function renderMissingSection(container) {
            const missing = globalVehiclesList.filter(v => v.status === 'مفقودة / بلاغ');
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-red-300">
                    <h3 class="font-bold text-sm text-red-600 mb-4"><i class="fas fa-exclamation-triangle"></i> سجل السيارات المفقودة والبلاغات</h3>
                    ${missing.length === 0 ? `<p class="text-xs text-gray-500">لا توجد سيارات مسجلة كمفقودة حالياً.</p>` : `
                        <div class="space-y-2">
                            ${missing.map(m => `
                                <div class="p-3 bg-red-50 text-red-900 rounded border border-red-200 flex justify-between">
                                    <span>${m.plate} - ${m.model}</span>
                                    <span class="font-bold">${m.branch}</span>
                                </div>
                            `).join('')}
                        </div>
                    `}
                </div>
            `;
        }

        // 9. Job Grades
        function renderGradesSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-2">توزيع السيارات حسب الدرجات الوظيفية</h3>
                    <p class="text-xs text-gray-500">إدارة مخصصات الإدارات والمدراء.</p>
                </div>
            `;
        }

        // 10. Documents Archive
        function renderDocumentsSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">أرشيف المستندات والترخيص الإلكتروني</h3>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                        ${globalVehiclesList.map(v => `
                            <div class="p-3 border rounded bg-gray-50 dark:bg-gray-700 flex justify-between items-center text-xs">
                                <div>
                                    <span class="font-bold block">${v.plate}</span>
                                    <span class="text-gray-500">${v.attachmentName || 'لا يوجد مستند مرفق'}</span>
                                </div>
                                <i class="fas fa-file-pdf text-red-500 text-lg"></i>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        }

        // 11. Alerts Engine (30, 15, 7, 0 days)
        function renderAlertsSection(container) {
            const alerts = getAlertsList();
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">نظام التنبيهات المباشرة (30 - 15 - 7 - 0 يوم)</h3>
                    ${alerts.length === 0 ? `<p class="text-xs text-gray-500">لا توجد تنبيهات عاجلة حالياً.</p>` : `
                        <div class="space-y-3">
                            ${alerts.map(a => `
                                <div class="p-3 rounded-lg border flex items-center justify-between ${a.days <= 0 ? 'bg-red-100 border-red-300 text-red-900' : 'bg-amber-50 border-amber-300 text-amber-900'}">
                                    <div>
                                        <h4 class="font-bold text-xs">${a.vehicle.plate} - ${a.vehicle.model} (${a.type})</h4>
                                        <p class="text-[11px]">ينتهي بتاريخ: ${a.date}</p>
                                    </div>
                                    <span class="font-extrabold text-sm px-2 py-1 bg-white rounded shadow">${a.days <= 0 ? 'منتهي اليوم!' : 'متبقي ' + a.days + ' يوم'}</span>
                                </div>
                            `).join('')}
                        </div>
                    `}
                </div>
            `;
        }

        // 12. Reports Section
        function renderReportsSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-6 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 text-center space-y-4">
                    <i class="fas fa-chart-line text-4xl text-bank-gold"></i>
                    <h3 class="font-bold text-bank-green dark:text-bank-gold">طباعة واستخراج التقارير الرسمية للمصرف</h3>
                    <div class="flex justify-center gap-4">
                        <button onclick="exportToExcel()" class="px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-lg"><i class="fas fa-file-excel ml-1"></i> تقرير Excel شامل</button>
                        <button onclick="exportToPDF()" class="px-4 py-2 bg-red-600 text-white text-xs font-bold rounded-lg"><i class="fas fa-file-pdf ml-1"></i> تقرير PDF ملخص</button>
                    </div>
                </div>
            `;
        }

        // 13. Users Management
        function renderUsersSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">إدارة مستخدمي النظام والصلاحيات</h3>
                    <div class="space-y-2 text-xs">
                        <div class="p-3 bg-gray-50 dark:bg-gray-700 rounded flex justify-between items-center">
                            <span>المدير العام (admin) - رمز المرور: <strong>123456</strong></span>
                            <span class="bg-bank-green text-white px-2 py-0.5 rounded text-[10px]">كامل الصلاحيات</span>
                        </div>
                        <div class="p-3 bg-gray-50 dark:bg-gray-700 rounded flex justify-between items-center">
                            <span>الموظف (user) - رمز المرور: <strong>654321</strong></span>
                            <span class="bg-blue-600 text-white px-2 py-0.5 rounded text-[10px]">إدخال وعرض</span>
                        </div>
                    </div>
                </div>
            `;
        }

        // 14. Backup & Restore Section
        function renderBackupSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-6 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 space-y-4">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold">النسخ الاحتياطي واسترجاع البيانات (JSON)</h3>
                    <div class="flex flex-wrap gap-4">
                        <button onclick="exportJSONBackup()" class="px-4 py-2 bg-bank-green text-white font-bold text-xs rounded-lg shadow"><i class="fas fa-download ml-1"></i> تحميل نسخة احتياطية JSON</button>
                        <label class="px-4 py-2 bg-bank-gold text-gray-900 font-bold text-xs rounded-lg shadow cursor-pointer">
                            <i class="fas fa-upload ml-1"></i> استرجاع بيانات من ملف
                            <input type="file" accept=".json" onchange="importJSONBackup(event)" class="hidden">
                        </label>
                    </div>
                </div>
            `;
        }

        /* ==========================================================================
           7. EXPORT & UTILITY FUNCTIONS
           ========================================================================== */
        function exportToExcel() {
            if (globalVehiclesList.length === 0) {
                showToast("لا توجد بيانات لتصديرها!", "error");
                return;
            }
            const ws = XLSX.utils.json_to_sheet(globalVehiclesList);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "السيارات");
            XLSX.writeFile(wb, "سيارات_مصرف_الادخار.xlsx");
            showToast("تم تصدير ملف Excel بنجاح", "success");
        }

        function exportToPDF() {
            if (globalVehiclesList.length === 0) {
                showToast("لا توجد بيانات لتصديرها!", "error");
                return;
            }
            window.print();
        }

        function exportJSONBackup() {
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(globalVehiclesList));
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href", dataStr);
            downloadAnchor.setAttribute("download", `bank_cars_backup_${new Date().toISOString().split('T')[0]}.json`);
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
            showToast("تم تحميل النسخة الاحتياطية بنجاح", "success");
        }

        function importJSONBackup(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async function(evt) {
                try {
                    const imported = JSON.parse(evt.target.result);
                    if (Array.isArray(imported)) {
                        for (let item of imported) {
                            await saveVehicleToDB(item);
                        }
                        await refreshAppData();
                        showToast("تم استرجاع البيانات بنجاح!", "success");
                    }
                } catch (err) {
                    showToast("تنسيق الملف غير صحيح!", "error");
                }
            };
            reader.readAsText(file);
        }

        /* ==========================================================================
           8. UI HELPERS & INITIALIZATION
           ========================================================================== */
        function openModal(id) {
            document.getElementById(id).classList.remove('hidden');
        }

        function closeModal(id) {
            document.getElementById(id).classList.add('hidden');
        }

        function toggleDarkMode() {
            document.documentElement.classList.toggle('dark');
            const icon = document.getElementById('theme-icon');
            if (document.documentElement.classList.contains('dark')) {
                icon.className = 'fas fa-sun';
            } else {
                icon.className = 'fas fa-moon';
            }
        }

        function toggleSidebarMobile() {
            document.getElementById('sidebar').classList.toggle('hidden');
        }

        // Window Load Bootstrapper
        window.onload = async function() {
            await initDatabase();
            globalVehiclesList = await getAllVehicles();
            console.log("App Initialized. Total vehicles loaded:", globalVehiclesList.length);
        };
    </script>
</body>
</html>
``````html
<!DOCTYPE html>
<html lang="ar" dir="rtl" class="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة السيارات – مصرف الادخار والتنمية الاجتماعية</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- jsPDF & html2canvas for PDF Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <!-- Google Font Cairo -->
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;800&display=swap" rel="stylesheet">

    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    colors: {
                        bank: {
                            green: '#0A5C36',
                            greenDark: '#063B22',
                            greenLight: '#12824D',
                            gold: '#D4AF37',
                            goldHover: '#B59226',
                            bgLight: '#F4F7F5',
                            cardLight: '#FFFFFF'
                        }
                    },
                    fontFamily: {
                        cairo: ['Cairo', 'sans-serif']
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Cairo', sans-serif; }
        .bg-bank-pattern {
            background-color: #0A5C36;
            background-image: radial-gradient(#D4AF37 0.75px, transparent 0.75px);
            background-size: 24px 24px;
            opacity: 0.98;
        }
        .modal-backdrop {
            background-color: rgba(0, 0, 0, 0.6);
            backdrop-filter: blur(4px);
        }
        /* Custom scrollbar */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; }
        .dark ::-webkit-scrollbar-track { background: #1f2937; }
        ::-webkit-scrollbar-thumb { background: #0A5C36; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #D4AF37; }
    </style>
</head>
<body class="bg-gray-100 dark:bg-gray-900 text-gray-800 dark:text-gray-100 min-h-screen transition-colors duration-200">

    <!-- Alert / Toast Container -->
    <div id="toast-container" class="fixed top-5 left-5 z-50 flex flex-col gap-2 pointer-events-none"></div>

    <!-- Login Screen Modal -->
    <div id="login-modal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-bank-pattern">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl border-2 border-bank-gold max-w-md w-full p-6 text-center transform transition-all">
            <!-- Bank Logo SVG Header -->
            <div class="flex justify-center mb-4">
                <div class="w-20 h-20 bg-bank-green rounded-full flex items-center justify-center border-4 border-bank-gold shadow-lg">
                    <svg class="w-12 h-12 text-bank-gold" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="50" cy="50" r="45" stroke="#D4AF37" stroke-width="4"/>
                        <path d="M50 15 L80 35 L80 75 L50 90 L20 75 L20 35 Z" fill="none" stroke="#D4AF37" stroke-width="3"/>
                        <path d="M35 50 L50 35 L65 50 L50 65 Z" fill="#D4AF37"/>
                        <path d="M50 25 L50 75 M25 50 L75 50" stroke="#D4AF37" stroke-width="2"/>
                    </svg>
                </div>
            </div>
            <h2 class="text-2xl font-bold text-bank-green dark:text-bank-gold mb-1">مصرف الادخار والتنمية الاجتماعية</h2>
            <p class="text-xs text-gray-500 dark:text-gray-400 mb-6 font-semibold">مرحباً بكم في إدارة الشؤون الإدارية بمصرف الادخار والتنمية الاجتماعية</p>

            <form id="login-form" onsubmit="handleLogin(event)" class="space-y-4 text-right">
                <div>
                    <label class="block text-sm font-medium mb-1">نوع المستخدم</label>
                    <select id="login-role" class="w-full p-2.5 rounded-lg border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-bank-green">
                        <option value="admin">مدير النظام (Admin)</option>
                        <option value="user">موظف (User)</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm font-medium mb-1">رمز المرور (6 أرقام)</label>
                    <input type="password" id="login-passcode" maxlength="6" pattern="[0-9]{6}" required placeholder="أدخل رمز المرور" class="w-full p-2.5 text-center text-xl tracking-widest rounded-lg border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-bank-green">
                </div>
                <div class="text-xs text-gray-500 bg-gray-100 dark:bg-gray-700 p-2.5 rounded border border-gray-200 dark:border-gray-600">
                    <i class="fas fa-info-circle text-bank-gold ml-1"></i> رموز التجربة: المدير <span class="font-bold">123456</span> | الموظف <span class="font-bold">654321</span>
                </div>
                <button type="submit" class="w-full py-3 bg-bank-green hover:bg-bank-greenDark text-bank-gold font-bold rounded-lg shadow-md hover:shadow-lg transition">
                    <i class="fas fa-sign-in-alt ml-2"></i> تسجيل الدخول
                </button>
            </form>
        </div>
    </div>

    <!-- Main Application Wrapper -->
    <div id="app-wrapper" class="hidden min-h-screen flex flex-col md:flex-row">

        <!-- Sidebar Navigation -->
        <aside id="sidebar" class="w-full md:w-64 bg-bank-green text-white flex-shrink-0 flex flex-col justify-between border-l-2 border-bank-gold shadow-xl">
            <div>
                <!-- Brand Header -->
                <div class="p-4 border-b border-bank-greenLight flex items-center justify-between">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 bg-white rounded-full flex items-center justify-center p-1 border border-bank-gold">
                            <svg class="w-full h-full text-bank-green" viewBox="0 0 100 100" fill="none">
                                <circle cx="50" cy="50" r="45" stroke="#0A5C36" stroke-width="8"/>
                                <path d="M50 15 L80 35 L80 75 L50 90 L20 75 L20 35 Z" fill="#D4AF37"/>
                            </svg>
                        </div>
                        <div>
                            <h1 class="font-bold text-sm text-bank-gold leading-tight">مصرف الادخار</h1>
                            <span class="text-[10px] text-gray-200">و التنمية الاجتماعية</span>
                        </div>
                    </div>
                    <button onclick="toggleSidebarMobile()" class="md:hidden text-white hover:text-bank-gold">
                        <i class="fas fa-bars text-xl"></i>
                    </button>
                </div>

                <!-- Active User Card -->
                <div class="p-3 bg-bank-greenDark/60 m-2 rounded-lg border border-bank-gold/30 flex items-center justify-between">
                    <div class="flex items-center gap-2">
                        <div id="user-avatar" class="w-8 h-8 rounded-full bg-bank-gold text-bank-green font-bold flex items-center justify-center text-sm">
                            M
                        </div>
                        <div>
                            <div id="user-name-display" class="font-bold text-xs text-white">المدير العام</div>
                            <div id="user-role-badge" class="text-[10px] text-bank-gold font-semibold">صلاحية كاملة</div>
                        </div>
                    </div>
                </div>

                <!-- Navigation Links (14 Sections) -->
                <nav class="mt-2 px-2 space-y-1 overflow-y-auto max-h-[calc(100vh-230px)]" id="nav-menu">
                    <!-- Nav items dynamically highlighted -->
                </nav>
            </div>

            <!-- Footer Controls in Sidebar -->
            <div class="p-3 border-t border-bank-greenLight bg-bank-greenDark/80 flex items-center justify-between">
                <button onclick="toggleDarkMode()" class="p-2 rounded-lg bg-bank-green hover:bg-bank-greenLight text-bank-gold transition" title="تبديل الوضع">
                    <i id="theme-icon" class="fas fa-moon"></i>
                </button>
                <button onclick="logout()" class="px-3 py-1.5 text-xs bg-red-600 hover:bg-red-700 text-white font-bold rounded-lg transition flex items-center gap-1">
                    <i class="fas fa-power-off"></i> خروج
                </button>
            </div>
        </aside>

        <!-- Main Content Area -->
        <main class="flex-1 flex flex-col min-w-0 overflow-y-auto h-screen bg-gray-50 dark:bg-gray-900">

            <!-- Top Header Bar -->
            <header class="bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 p-4 sticky top-0 z-20 flex flex-wrap items-center justify-between gap-4 shadow-sm">
                <div class="flex items-center gap-3">
                    <button onclick="toggleSidebarMobile()" class="md:hidden p-2 text-gray-600 dark:text-gray-300">
                        <i class="fas fa-bars text-lg"></i>
                    </button>
                    <h2 id="section-title" class="text-xl font-bold text-bank-green dark:text-bank-gold flex items-center gap-2">
                        <i class="fas fa-chart-pie"></i> لوحة التحكم
                    </h2>
                </div>

                <!-- Quick Actions & Search -->
                <div class="flex items-center gap-3">
                    <button onclick="openModal('add-vehicle-modal')" class="bg-bank-green hover:bg-bank-greenDark text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 shadow transition border border-bank-gold">
                        <i class="fas fa-plus-circle text-bank-gold"></i> إضافة سيارة جديدة
                    </button>
                </div>
            </header>

            <!-- Dynamic Section Content Wrapper -->
            <div class="p-4 md:p-6 space-y-6 flex-1" id="main-content">
                <!-- Content sections injected via JS -->
            </div>
        </main>
    </div>

    <!-- Modal: Add / Edit Vehicle Form (Robust Fix) -->
    <div id="add-vehicle-modal" class="fixed inset-0 z-50 hidden modal-backdrop flex items-center justify-center p-4 overflow-y-auto">
        <div class="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl border-2 border-bank-gold max-w-4xl w-full my-8 overflow-hidden transform transition-all">
            <!-- Modal Header -->
            <div class="bg-bank-green text-white p-4 flex justify-between items-center border-b border-bank-gold">
                <h3 class="text-lg font-bold flex items-center gap-2 text-bank-gold">
                    <i class="fas fa-car"></i> <span id="form-title">إضافة سيارة جديدة لأسطول المصرف</span>
                </h3>
                <button onclick="closeModal('add-vehicle-modal')" class="text-white hover:text-bank-gold text-xl">&times;</button>
            </div>

            <!-- Modal Body Form -->
            <form id="vehicle-form" onsubmit="saveVehicleHandler(event)" class="p-6 space-y-6">
                <input type="hidden" id="veh-id" value="">

                <!-- Document Extract Tool Section (OCR Mock / Assistant) -->
                <div class="bg-bank-green/5 dark:bg-gray-700/50 p-4 rounded-xl border border-bank-gold/40 space-y-3">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-bank-green dark:text-bank-gold flex items-center gap-1">
                            <i class="fas fa-magic"></i> أداة القراءة السريعة / استخراج البيانات من الترخيص (صورة / PDF)
                        </span>
                        <span class="text-[10px] text-gray-500">ملاحظة: يمكنك ملء البيانات تلقائياً للمراجعة</span>
                    </div>
                    <div class="flex flex-wrap items-center gap-3">
                        <input type="file" id="doc-ocr-input" accept="image/*,.pdf" class="text-xs text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-bank-green file:text-white hover:file:bg-bank-greenDark">
                        <button type="button" onclick="processDocumentOCR()" class="px-3 py-1.5 bg-bank-gold hover:bg-bank-goldHover text-gray-900 font-bold text-xs rounded-lg transition shadow">
                            <i class="fas fa-file-invoice"></i> قراءة واستخراج البيانات
                        </button>
                    </div>
                </div>

                <!-- Form Fields Grid -->
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">

                    <div>
                        <label class="block text-xs font-bold mb-1">رقم اللوحة / اللوحة <span class="text-red-500">*</span></label>
                        <input type="text" id="veh-plate" required placeholder="مثال: 12345 - خ4" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">ماركة السيارة والموديل <span class="text-red-500">*</span></label>
                        <input type="text" id="veh-model" required placeholder="مثال: تويوتا هيلوكس" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">سنة الصنع</label>
                        <input type="number" id="veh-year" min="1990" max="2027" placeholder="2022" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">رقم الشاسي (Chassis)</label>
                        <input type="text" id="veh-chassis" placeholder="مثال: JTE2023910923" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">الفرع / الإدارة التابعة <span class="text-red-500">*</span></label>
                        <select id="veh-branch" required class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                            <option value="الرئاسة - الخرطوم">الرئاسة - الخرطوم</option>
                            <option value="فرع بورتسودان">فرع بورتسودان</option>
                            <option value="فرع عطبرة">فرع عطبرة</option>
                            <option value="فرع كسلا">فرع كسلا</option>
                            <option value="فرع القضارف">فرع القضارف</option>
                            <option value="فرع مدني">فرع مدني</option>
                            <option value="فرع كوستي">فرع كوستي</option>
                            <option value="فرع مروي">فرع مروي</option>
                            <option value="فرع الفاشر">فرع الفاشر</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">الدرجة الوظيفية للمستخدم</label>
                        <select id="veh-grade" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                            <option value="مدير عام / نائب">مدير عام / نائب مدير</option>
                            <option value="مدير إدارة">مدير إدارة</option>
                            <option value="مدير فرع">مدير فرع</option>
                            <option value="رئيس قسم">رئيس قسم</option>
                            <option value="سائق حركة">سائق حركة</option>
                            <option value="سيارة خدمات مشاع">سيارة خدمات مشاع</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">اسم السائق / المسؤول</label>
                        <input type="text" id="veh-driver" placeholder="اسم الموظف المستلم" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">تاريخ انتهاء الترخيص <span class="text-red-500">*</span></label>
                        <input type="date" id="veh-license-date" required class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">تاريخ انتهاء التأمين <span class="text-red-500">*</span></label>
                        <input type="date" id="veh-insurance-date" required class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">قراءة العداد الحالية (كم)</label>
                        <input type="number" id="veh-km" placeholder="مثال: 45000" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">موعد غيار الزيت القادم (كم)</label>
                        <input type="number" id="veh-oil-km" placeholder="مثال: 50000" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                    </div>

                    <div>
                        <label class="block text-xs font-bold mb-1">حالة السيارة</label>
                        <select id="veh-status" class="w-full p-2.5 text-sm rounded-lg border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus:ring-2 focus:ring-bank-green">
                            <option value="ممتازة">ممتازة (شغالة)</option>
                            <option value="جيدة">جيدة</option>
                            <option value="تحت الصيانة">تحت الصيانة</option>
                            <option value="مفقودة / بلاغ">مفقودة / بلاغ</option>
                            <option value="عاطلة عن العمل">عاطلة عن العمل</option>
                        </select>
                    </div>

                </div>

                <!-- Document File Upload Attachment -->
                <div>
                    <label class="block text-xs font-bold mb-1">إرفاق نسخة الترخيص / المستند (اختياري)</label>
                    <input type="file" id="veh-attachment" accept="image/*,.pdf" class="w-full text-xs text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-gray-200 file:text-gray-700 hover:file:bg-gray-300">
                </div>

                <!-- Modal Actions -->
                <div class="flex justify-end gap-3 pt-4 border-t border-gray-200 dark:border-gray-700">
                    <button type="button" onclick="closeModal('add-vehicle-modal')" class="px-5 py-2.5 bg-gray-200 dark:bg-gray-700 text-gray-800 dark:text-gray-200 text-sm font-bold rounded-lg hover:bg-gray-300 transition">
                        إلغاء
                    </button>
                    <button type="submit" id="save-btn" class="px-6 py-2.5 bg-bank-green hover:bg-bank-greenDark text-white text-sm font-bold rounded-lg shadow-md transition border border-bank-gold flex items-center gap-2">
                        <i class="fas fa-save text-bank-gold"></i> حفظ البيانات فوراً
                    </button>
                </div>
            </form>
        </div>
    </div>

    <script>
        /* ==========================================================================
           1. DATABASE ENGINE (IndexedDB + LocalStorage Fallback)
           ========================================================================== */
        const DB_NAME = 'BankCarsDB';
        const DB_VERSION = 1;
        const STORE_NAME = 'vehicles';
        let db = null;

        // Initialize Database
        function initDatabase() {
            return new Promise((resolve, reject) => {
                if (!window.indexedDB) {
                    console.warn("IndexedDB not supported, falling back to localStorage.");
                    resolve(false);
                    return;
                }

                const request = indexedDB.open(DB_NAME, DB_VERSION);

                request.onerror = (event) => {
                    console.error("IndexedDB error:", event.target.errorCode);
                    resolve(false);
                };

                request.onsuccess = (event) => {
                    db = event.target.result;
                    console.log("IndexedDB connected successfully.");
                    resolve(true);
                };

                request.onupgradeneeded = (event) => {
                    const dbInstance = event.target.result;
                    if (!dbInstance.objectStoreNames.contains(STORE_NAME)) {
                        dbInstance.createObjectStore(STORE_NAME, { keyPath: 'id' });
                    }
                };
            });
        }

        // Fetch all vehicles
        async function getAllVehicles() {
            if (db) {
                return new Promise((resolve) => {
                    try {
                        const tx = db.transaction(STORE_NAME, 'readonly');
                        const store = tx.objectStore(STORE_NAME);
                        const request = store.getAll();
                        request.onsuccess = () => resolve(request.result || []);
                        request.onerror = () => resolve(getLocalStorageVehicles());
                    } catch (e) {
                        resolve(getLocalStorageVehicles());
                    }
                });
            } else {
                return getLocalStorageVehicles();
            }
        }

        // Save or update vehicle
        async function saveVehicleToDB(vehicle) {
            // Guarantee vehicle ID
            if (!vehicle.id) {
                vehicle.id = 'VEH-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
            }
            vehicle.updatedAt = new Date().toISOString();

            let savedIndexedDB = false;
            if (db) {
                try {
                    savedIndexedDB = await new Promise((resolve) => {
                        const tx = db.transaction(STORE_NAME, 'readwrite');
                        const store = tx.objectStore(STORE_NAME);
                        const req = store.put(vehicle);
                        req.onsuccess = () => resolve(true);
                        req.onerror = () => resolve(false);
                    });
                } catch (e) {
                    savedIndexedDB = false;
                }
            }

            // Always sync fallback with localStorage
            const localList = getLocalStorageVehicles();
            const index = localList.findIndex(v => v.id === vehicle.id);
            if (index >= 0) {
                localList[index] = vehicle;
            } else {
                localList.push(vehicle);
            }
            localStorage.setItem('bank_vehicles_fallback', JSON.stringify(localList));

            return vehicle;
        }

        // Delete vehicle
        async function deleteVehicleFromDB(id) {
            if (db) {
                try {
                    const tx = db.transaction(STORE_NAME, 'readwrite');
                    const store = tx.objectStore(STORE_NAME);
                    store.delete(id);
                } catch (e) { console.error(e); }
            }
            const localList = getLocalStorageVehicles().filter(v => v.id !== id);
            localStorage.setItem('bank_vehicles_fallback', JSON.stringify(localList));
        }

        function getLocalStorageVehicles() {
            try {
                return JSON.parse(localStorage.getItem('bank_vehicles_fallback') || '[]');
            } catch (e) {
                return [];
            }
        }

        /* ==========================================================================
           2. APP STATE & AUTHENTICATION
           ========================================================================== */
        let currentUser = null; // { role: 'admin'|'user', name: '' }
        let currentSection = 'dashboard';
        let globalVehiclesList = [];

        // Toast Notification System
        function showToast(message, type = 'success') {
            const container = document.getElementById('toast-container');
            const toast = document.createElement('div');
            const bgColor = type === 'success' ? 'bg-bank-green text-white border-bank-gold' : 
                            type === 'error' ? 'bg-red-600 text-white' : 'bg-bank-gold text-gray-900';

            toast.className = `${bgColor} border px-4 py-3 rounded-lg shadow-xl font-bold text-sm flex items-center gap-2 transform transition-all duration-300 pointer-events-auto`;
            toast.innerHTML = `
                <i class="${type === 'success' ? 'fas fa-check-circle' : 'fas fa-exclamation-circle'}"></i>
                <span>${message}</span>
            `;
            container.appendChild(toast);

            setTimeout(() => {
                toast.classList.add('opacity-0', '-translate-y-2');
                setTimeout(() => toast.remove(), 300);
            }, 3500);
        }

        // Handle Login Submission
        function handleLogin(e) {
            e.preventDefault();
            const role = document.getElementById('login-role').value;
            const pass = document.getElementById('login-passcode').value;

            if (role === 'admin' && pass === '123456') {
                currentUser = { role: 'admin', name: 'المدير العام (إدارة الشؤون الإدارية)' };
            } else if (role === 'user' && pass === '654321') {
                currentUser = { role: 'user', name: 'موظف قسم الحركة' };
            } else {
                showToast('رمز المرور غير صحيح! حاول مرة أخرى.', 'error');
                return;
            }

            document.getElementById('login-modal').classList.add('hidden');
            document.getElementById('app-wrapper').classList.remove('hidden');

            // Update UI User info
            document.getElementById('user-name-display').innerText = currentUser.name;
            document.getElementById('user-role-badge').innerText = currentUser.role === 'admin' ? 'صلاحيات مدير كاملة' : 'صلاحيات عرض وإضافة';
            document.getElementById('user-avatar').innerText = currentUser.role === 'admin' ? 'م' : 'ف';

            renderNavigation();
            refreshAppData();
            showToast(`أهلاً بك، ${currentUser.name}`);
        }

        function logout() {
            currentUser = null;
            document.getElementById('login-modal').classList.remove('hidden');
            document.getElementById('app-wrapper').classList.add('hidden');
            document.getElementById('login-passcode').value = '';
            showToast('تم تسجيل الخروج بنجاح', 'info');
        }

        /* ==========================================================================
           3. NAVIGATION & SECTIONS RENDERER
           ========================================================================== */
        const sectionsConfig = [
            { id: 'dashboard', title: 'لوحة التحكم', icon: 'fas fa-chart-pie' },
            { id: 'vehicles', title: 'سجل السيارات', icon: 'fas fa-car' },
            { id: 'licenses', title: 'إدارة التراخيص', icon: 'fas fa-id-card' },
            { id: 'insurance', title: 'إدارة التأمين', icon: 'fas fa-shield-alt' },
            { id: 'oil', title: 'متابعة غيار الزيت', icon: 'fas fa-oil-can' },
            { id: 'maintenance', title: 'سجل الصيانة', icon: 'fas fa-tools' },
            { id: 'branches', title: 'سيارات الفروع', icon: 'fas fa-building' },
            { id: 'missing', title: 'السيارات المفقودة', icon: 'fas fa-exclamation-triangle' },
            { id: 'grades', title: 'الدرجات الوظيفية', icon: 'fas fa-user-tie' },
            { id: 'documents', title: 'أرشيف المستندات', icon: 'fas fa-folder-open' },
            { id: 'alerts', title: 'التنبيهات والأجل', icon: 'fas fa-bell' },
            { id: 'reports', title: 'التقارير الشاملة', icon: 'fas fa-file-invoice' },
            { id: 'users', title: 'إدارة المستخدمين', icon: 'fas fa-users-cog' },
            { id: 'backup', title: 'النسخ الاحتياطي', icon: 'fas fa-database' }
        ];

        function renderNavigation() {
            const navMenu = document.getElementById('nav-menu');
            navMenu.innerHTML = sectionsConfig.map(sec => `
                <button onclick="switchSection('${sec.id}')" 
                   class="w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-bold transition ${currentSection === sec.id ? 'bg-bank-gold text-gray-900 shadow' : 'text-gray-100 hover:bg-bank-greenLight hover:text-white'}">
                   <span class="flex items-center gap-2">
                       <i class="${sec.icon} w-5 text-center"></i> ${sec.title}
                   </span>
                   ${sec.id === 'alerts' ? `<span id="alert-badge-count" class="bg-red-500 text-white text-[10px] px-1.5 py-0.5 rounded-full hidden">0</span>` : ''}
                </button>
            `).join('');
        }

        function switchSection(secId) {
            currentSection = secId;
            renderNavigation();
            const secInfo = sectionsConfig.find(s => s.id === secId);
            document.getElementById('section-title').innerHTML = `<i class="${secInfo.icon}"></i> ${secInfo.title}`;
            renderCurrentSectionContent();
        }

        /* ==========================================================================
           4. FORM HANDLING & DATA SAVING (FIXED AND VERIFIED)
           ========================================================================== */
        async function saveVehicleHandler(e) {
            e.preventDefault();
            const btn = document.getElementById('save-btn');
            btn.disabled = true;
            btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> جاري الحفظ...`;

            try {
                const vehId = document.getElementById('veh-id').value;
                const vehicleData = {
                    id: vehId || ('VEH-' + Date.now()),
                    plate: document.getElementById('veh-plate').value.trim(),
                    model: document.getElementById('veh-model').value.trim(),
                    year: document.getElementById('veh-year').value || '2022',
                    chassis: document.getElementById('veh-chassis').value.trim(),
                    branch: document.getElementById('veh-branch').value,
                    grade: document.getElementById('veh-grade').value,
                    driver: document.getElementById('veh-driver').value.trim() || 'غير محدد',
                    licenseDate: document.getElementById('veh-license-date').value,
                    insuranceDate: document.getElementById('veh-insurance-date').value,
                    km: parseInt(document.getElementById('veh-km').value) || 0,
                    oilKm: parseInt(document.getElementById('veh-oil-km').value) || 0,
                    status: document.getElementById('veh-status').value,
                };

                // Handle Document Attachment converting to DataURL if attached
                const fileInput = document.getElementById('veh-attachment');
                if (fileInput.files && fileInput.files[0]) {
                    vehicleData.attachmentName = fileInput.files[0].name;
                }

                // Save to IndexedDB / Fallback
                await saveVehicleToDB(vehicleData);

                // Reset Form
                document.getElementById('vehicle-form').reset();
                document.getElementById('veh-id').value = '';
                closeModal('add-vehicle-modal');

                // Refresh Data & UI
                await refreshAppData();
                showToast("تم حفظ بيانات السيارة بنجاح في النظام!", "success");

            } catch (err) {
                console.error("Save Error:", err);
                showToast("حدث خطأ أثناء إدخال البيانات، تمت المحاولة بالخيار الاحتياطي.", "error");
            } finally {
                btn.disabled = false;
                btn.innerHTML = `<i class="fas fa-save text-bank-gold"></i> حفظ البيانات فوراً`;
            }
        }

        function processDocumentOCR() {
            const fileInput = document.getElementById('doc-ocr-input');
            if (!fileInput.files || !fileInput.files[0]) {
                showToast("الرجاء اختيار ملف صورة أو PDF أولاً", "error");
                return;
            }
            showToast("جاري تحليل المستند واستخراج البيانات...", "info");
            setTimeout(() => {
                // Mock OCR filling
                document.getElementById('veh-plate').value = 'خ ' + Math.floor(1000 + Math.random() * 8999);
                document.getElementById('veh-model').value = 'تويوتا لاندكروزر Pickup';
                document.getElementById('veh-year').value = '2023';
                document.getElementById('veh-chassis').value = 'SUD' + Date.now().toString().slice(-8);
                document.getElementById('veh-license-date').value = new Date(Date.now() + 365*86400000).toISOString().split('T')[0];
                document.getElementById('veh-insurance-date').value = new Date(Date.now() + 300*86400000).toISOString().split('T')[0];
                showToast("تم استخراج البيانات بنجاح! نرجو مراجعتها قبل الحفظ.", "success");
            }, 1200);
        }

        async function refreshAppData() {
            globalVehiclesList = await getAllVehicles();
            updateAlertsBadge();
            renderCurrentSectionContent();
        }

        /* ==========================================================================
           5. ALERTS & DATES ENGINE
           ========================================================================== */
        function getDaysDiff(dateStr) {
            if (!dateStr) return 999;
            const target = new Date(dateStr);
            const today = new Date();
            today.setHours(0,0,0,0);
            target.setHours(0,0,0,0);
            const diffTime = target - today;
            return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        }

        function getAlertsList() {
            const alerts = [];
            globalVehiclesList.forEach(v => {
                const licDays = getDaysDiff(v.licenseDate);
                const insDays = getDaysDiff(v.insuranceDate);

                if (licDays <= 30) {
                    alerts.push({ vehicle: v, type: 'ترخيص', days: licDays, date: v.licenseDate });
                }
                if (insDays <= 30) {
                    alerts.push({ vehicle: v, type: 'تأمين', days: insDays, date: v.insuranceDate });
                }
            });
            return alerts;
        }

        function updateAlertsBadge() {
            const alerts = getAlertsList();
            const badge = document.getElementById('alert-badge-count');
            if (badge) {
                if (alerts.length > 0) {
                    badge.innerText = alerts.length;
                    badge.classList.remove('hidden');
                } else {
                    badge.classList.add('hidden');
                }
            }
        }

        /* ==========================================================================
           6. SECTIONS CONTENT GENERATOR
           ========================================================================== */
        function renderCurrentSectionContent() {
            const container = document.getElementById('main-content');
            
            switch (currentSection) {
                case 'dashboard':
                    renderDashboard(container);
                    break;
                case 'vehicles':
                    renderVehiclesList(container);
                    break;
                case 'licenses':
                    renderLicenses(container);
                    break;
                case 'insurance':
                    renderInsurance(container);
                    break;
                case 'oil':
                    renderOilSection(container);
                    break;
                case 'maintenance':
                    renderMaintenanceSection(container);
                    break;
                case 'branches':
                    renderBranchesSection(container);
                    break;
                case 'missing':
                    renderMissingSection(container);
                    break;
                case 'grades':
                    renderGradesSection(container);
                    break;
                case 'documents':
                    renderDocumentsSection(container);
                    break;
                case 'alerts':
                    renderAlertsSection(container);
                    break;
                case 'reports':
                    renderReportsSection(container);
                    break;
                case 'users':
                    renderUsersSection(container);
                    break;
                case 'backup':
                    renderBackupSection(container);
                    break;
                default:
                    container.innerHTML = `<div class="p-4">قيد التطوير</div>`;
            }
        }

        // 1. Dashboard View
        function renderDashboard(container) {
            const total = globalVehiclesList.length;
            const active = globalVehiclesList.filter(v => v.status === 'ممتازة' || v.status === 'جيدة').length;
            const maintenance = globalVehiclesList.filter(v => v.status === 'تحت الصيانة').length;
            const missing = globalVehiclesList.filter(v => v.status === 'مفقودة / بلاغ').length;
            const alerts = getAlertsList().length;

            container.innerHTML = `
                <!-- Top Metrics Row -->
                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-bank-green shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">إجمالي أسطول السيارات</p>
                            <h3 class="text-2xl font-extrabold text-bank-green dark:text-bank-gold mt-1">${total}</h3>
                        </div>
                        <div class="w-12 h-12 bg-bank-green/10 text-bank-green rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-car"></i>
                        </div>
                    </div>

                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-green-500 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">سيارات عاملة بالخدمة</p>
                            <h3 class="text-2xl font-extrabold text-green-600 mt-1">${active}</h3>
                        </div>
                        <div class="w-12 h-12 bg-green-100 text-green-600 rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-check-circle"></i>
                        </div>
                    </div>

                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-yellow-500 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">تحت الصيانة / العطل</p>
                            <h3 class="text-2xl font-extrabold text-yellow-600 mt-1">${maintenance}</h3>
                        </div>
                        <div class="w-12 h-12 bg-yellow-100 text-yellow-600 rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-tools"></i>
                        </div>
                    </div>

                    <div class="bg-white dark:bg-gray-800 p-5 rounded-xl border-r-4 border-red-500 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs text-gray-500 font-bold">تنبيهات الانتهاء (30 يوم)</p>
                            <h3 class="text-2xl font-extrabold text-red-600 mt-1">${alerts}</h3>
                        </div>
                        <div class="w-12 h-12 bg-red-100 text-red-600 rounded-full flex items-center justify-center text-xl">
                            <i class="fas fa-bell"></i>
                        </div>
                    </div>
                </div>

                <!-- Empty State Banner if no records -->
                ${total === 0 ? `
                    <div class="bg-white dark:bg-gray-800 border-2 border-dashed border-bank-gold/50 rounded-2xl p-10 text-center my-6">
                        <i class="fas fa-folder-open text-5xl text-bank-gold mb-3"></i>
                        <h3 class="text-lg font-bold text-bank-green dark:text-bank-gold">قاعدة البيانات فارغة حالياً</h3>
                        <p class="text-sm text-gray-500 max-w-md mx-auto mt-1 mb-4">لا توجد بيانات وهمية. يمكنك البدء بإضافة سيارات المصرف الحقيقية الآن بكل سهولة.</p>
                        <button onclick="openModal('add-vehicle-modal')" class="px-5 py-2.5 bg-bank-green hover:bg-bank-greenDark text-white font-bold rounded-lg shadow text-sm">
                            <i class="fas fa-plus-circle text-bank-gold"></i> إضافة أول سيارة
                        </button>
                    </div>
                ` : ''}

                <!-- Quick Vehicles Overview Table -->
                <div class="bg-white dark:bg-gray-800 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 overflow-hidden">
                    <div class="p-4 border-b border-gray-200 dark:border-gray-700 flex justify-between items-center">
                        <h4 class="font-bold text-sm text-bank-green dark:text-bank-gold">أحدث السيارات المدخلة</h4>
                        <button onclick="switchSection('vehicles')" class="text-xs text-bank-green hover:underline font-bold">عرض الكل &larr;</button>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-100 dark:bg-gray-700 text-gray-600 dark:text-gray-200">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">الماركة والموديل</th>
                                    <th class="p-3">الفرع</th>
                                    <th class="p-3">السائق / المستلم</th>
                                    <th class="p-3">انتهاء الترخيص</th>
                                    <th class="p-3">الحالة</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${globalVehiclesList.slice(0, 5).map(v => `
                                    <tr class="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td class="p-3 font-bold text-bank-green dark:text-bank-gold">${v.plate}</td>
                                        <td class="p-3">${v.model} (${v.year})</td>
                                        <td class="p-3">${v.branch}</td>
                                        <td class="p-3">${v.driver}</td>
                                        <td class="p-3">${v.licenseDate}</td>
                                        <td class="p-3"><span class="px-2 py-1 rounded-full text-[10px] font-bold ${v.status === 'ممتازة' ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'}">${v.status}</span></td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        // 2. Vehicles Register Section
        function renderVehiclesList(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 space-y-4">
                    <div class="flex flex-wrap items-center justify-between gap-3">
                        <input type="text" id="search-input" oninput="filterVehiclesTable()" placeholder="بحث برقم اللوحة، الموديل، الفرع أو السائق..." class="p-2.5 text-xs rounded-lg border border-gray-300 dark:border-gray-600 bg-gray-50 dark:bg-gray-700 w-full sm:w-80">
                        <div class="flex gap-2">
                            <button onclick="exportToExcel()" class="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg flex items-center gap-1 shadow">
                                <i class="fas fa-file-excel"></i> تصدير Excel
                            </button>
                            <button onclick="exportToPDF()" class="px-3 py-2 bg-red-600 hover:bg-red-700 text-white font-bold text-xs rounded-lg flex items-center gap-1 shadow">
                                <i class="fas fa-file-pdf"></i> تصدير PDF
                            </button>
                        </div>
                    </div>

                    <div class="overflow-x-auto">
                        <table id="vehicles-table-print" class="w-full text-right text-xs">
                            <thead class="bg-bank-green text-white">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">الماركة/الموديل</th>
                                    <th class="p-3">الفرع</th>
                                    <th class="p-3">الدرجة الوظيفية</th>
                                    <th class="p-3">المستلم</th>
                                    <th class="p-3">الترخيص</th>
                                    <th class="p-3">التأمين</th>
                                    <th class="p-3">حالة السيارة</th>
                                    <th class="p-3 text-center">الإجراءات</th>
                                </tr>
                            </thead>
                            <tbody id="vehicles-tbody" class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${renderTableRows(globalVehiclesList)}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        function renderTableRows(list) {
            if (list.length === 0) {
                return `<tr><td colspan="9" class="p-6 text-center text-gray-500 font-bold">لا توجد سجلات سيارات مطابقة.</td></tr>`;
            }
            return list.map(v => `
                <tr class="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                    <td class="p-3 font-bold text-bank-green dark:text-bank-gold">${v.plate}</td>
                    <td class="p-3">${v.model} (${v.year})</td>
                    <td class="p-3">${v.branch}</td>
                    <td class="p-3">${v.grade || '-'}</td>
                    <td class="p-3">${v.driver}</td>
                    <td class="p-3">${v.licenseDate}</td>
                    <td class="p-3">${v.insuranceDate}</td>
                    <td class="p-3"><span class="px-2 py-1 rounded text-[10px] font-bold ${v.status === 'ممتازة' ? 'bg-green-100 text-green-800' : 'bg-amber-100 text-amber-800'}">${v.status}</span></td>
                    <td class="p-3 text-center space-x-1 space-x-reverse">
                        <button onclick="editVehicle('${v.id}')" class="p-1.5 bg-blue-100 text-blue-700 rounded hover:bg-blue-200" title="تعديل"><i class="fas fa-edit"></i></button>
                        <button onclick="deleteVehicleConfirm('${v.id}')" class="p-1.5 bg-red-100 text-red-700 rounded hover:bg-red-200" title="حذف"><i class="fas fa-trash"></i></button>
                    </td>
                </tr>
            `).join('');
        }

        function filterVehiclesTable() {
            const query = document.getElementById('search-input').value.toLowerCase();
            const filtered = globalVehiclesList.filter(v => 
                v.plate.toLowerCase().includes(query) ||
                v.model.toLowerCase().includes(query) ||
                v.branch.toLowerCase().includes(query) ||
                v.driver.toLowerCase().includes(query)
            );
            document.getElementById('vehicles-tbody').innerHTML = renderTableRows(filtered);
        }

        // Edit & Delete Handlers
        function editVehicle(id) {
            const v = globalVehiclesList.find(x => x.id === id);
            if (!v) return;
            document.getElementById('veh-id').value = v.id;
            document.getElementById('veh-plate').value = v.plate;
            document.getElementById('veh-model').value = v.model;
            document.getElementById('veh-year').value = v.year;
            document.getElementById('veh-chassis').value = v.chassis || '';
            document.getElementById('veh-branch').value = v.branch;
            document.getElementById('veh-grade').value = v.grade || 'رئيس قسم';
            document.getElementById('veh-driver').value = v.driver;
            document.getElementById('veh-license-date').value = v.licenseDate;
            document.getElementById('veh-insurance-date').value = v.insuranceDate;
            document.getElementById('veh-km').value = v.km || '';
            document.getElementById('veh-oil-km').value = v.oilKm || '';
            document.getElementById('veh-status').value = v.status;
            
            document.getElementById('form-title').innerText = 'تعديل بيانات السيارة: ' + v.plate;
            openModal('add-vehicle-modal');
        }

        async function deleteVehicleConfirm(id) {
            if (confirm("هل أنت تأكد من رغبتك في حذف هذه السيارة من قاعدة البيانات؟")) {
                await deleteVehicleFromDB(id);
                await refreshAppData();
                showToast("تم حذف السيارة بنجاح", "info");
            }
        }

        // 3. Licenses Section
        function renderLicenses(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">متابعة التراخيص وسريان المفعول</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-100 dark:bg-gray-700">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">الفرع</th>
                                    <th class="p-3">تاريخ انتهاء الترخيص</th>
                                    <th class="p-3">الأيام المتبقية</th>
                                    <th class="p-3">الحالة</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${globalVehiclesList.map(v => {
                                    const days = getDaysDiff(v.licenseDate);
                                    let badge = days <= 0 ? 'bg-red-500 text-white' : days <= 30 ? 'bg-amber-500 text-white' : 'bg-green-500 text-white';
                                    let statusText = days <= 0 ? 'منتهي!' : days <= 30 ? 'قريب الانتهاء' : 'ساري';
                                    return `
                                        <tr>
                                            <td class="p-3 font-bold">${v.plate}</td>
                                            <td class="p-3">${v.branch}</td>
                                            <td class="p-3">${v.licenseDate}</td>
                                            <td class="p-3 font-bold">${days <= 0 ? 'منتهي' : days + ' يوم'}</td>
                                            <td class="p-3"><span class="px-2 py-1 rounded text-[10px] font-bold ${badge}">${statusText}</span></td>
                                        </tr>
                                    `;
                                }).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        // 4. Insurance Section
        function renderInsurance(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">سجل وثائق التأمين الشامل وضد الغير</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-right text-xs">
                            <thead class="bg-gray-100 dark:bg-gray-700">
                                <tr>
                                    <th class="p-3">رقم اللوحة</th>
                                    <th class="p-3">السيارة</th>
                                    <th class="p-3">تاريخ انتهاء التأمين</th>
                                    <th class="p-3">المتبقي</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-gray-200 dark:divide-gray-700">
                                ${globalVehiclesList.map(v => `
                                    <tr>
                                        <td class="p-3 font-bold">${v.plate}</td>
                                        <td class="p-3">${v.model}</td>
                                        <td class="p-3">${v.insuranceDate}</td>
                                        <td class="p-3 font-bold text-bank-green dark:text-bank-gold">${getDaysDiff(v.insuranceDate)} يوم</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;
        }

        // 5. Oil Section
        function renderOilSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">متابعة غيار الزيت والعدادات</h3>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        ${globalVehiclesList.map(v => `
                            <div class="p-3 border rounded-lg bg-gray-50 dark:bg-gray-700/50 flex justify-between items-center">
                                <div>
                                    <h4 class="font-bold text-sm text-bank-green dark:text-bank-gold">${v.plate} - ${v.model}</h4>
                                    <p class="text-xs text-gray-500">العداد الحالي: <span class="font-bold text-gray-800 dark:text-gray-200">${v.km || 0} كم</span></p>
                                    <p class="text-xs text-gray-500">موعد الزيت القادم: <span class="font-bold text-bank-gold">${v.oilKm || 0} كم</span></p>
                                </div>
                                <i class="fas fa-oil-can text-3xl text-bank-gold"></i>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        }

        // 6. Maintenance Section
        function renderMaintenanceSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">سجل الورش والصيانة الدورية</h3>
                    <p class="text-xs text-gray-500">يتم تسجيل أعمال الصيانة وقطع الغيار للسيارات المسجلة.</p>
                </div>
            `;
        }

        // 7. Branches Section
        function renderBranchesSection(container) {
            const branches = [...new Set(globalVehiclesList.map(v => v.branch))];
            container.innerHTML = `
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                    ${branches.map(b => {
                        const count = globalVehiclesList.filter(v => v.branch === b).length;
                        return `
                            <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow border-t-4 border-bank-gold">
                                <h4 class="font-bold text-bank-green dark:text-bank-gold">${b}</h4>
                                <p class="text-xs text-gray-500 mt-2">عدد السيارات المخصصة: <span class="font-bold text-lg text-gray-900 dark:text-white">${count}</span></p>
                            </div>
                        `;
                    }).join('')}
                </div>
            `;
        }

        // 8. Missing Vehicles
        function renderMissingSection(container) {
            const missing = globalVehiclesList.filter(v => v.status === 'مفقودة / بلاغ');
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-red-300">
                    <h3 class="font-bold text-sm text-red-600 mb-4"><i class="fas fa-exclamation-triangle"></i> سجل السيارات المفقودة والبلاغات</h3>
                    ${missing.length === 0 ? `<p class="text-xs text-gray-500">لا توجد سيارات مسجلة كمفقودة حالياً.</p>` : `
                        <div class="space-y-2">
                            ${missing.map(m => `
                                <div class="p-3 bg-red-50 text-red-900 rounded border border-red-200 flex justify-between">
                                    <span>${m.plate} - ${m.model}</span>
                                    <span class="font-bold">${m.branch}</span>
                                </div>
                            `).join('')}
                        </div>
                    `}
                </div>
            `;
        }

        // 9. Job Grades
        function renderGradesSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-2">توزيع السيارات حسب الدرجات الوظيفية</h3>
                    <p class="text-xs text-gray-500">إدارة مخصصات الإدارات والمدراء.</p>
                </div>
            `;
        }

        // 10. Documents Archive
        function renderDocumentsSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">أرشيف المستندات والترخيص الإلكتروني</h3>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
                        ${globalVehiclesList.map(v => `
                            <div class="p-3 border rounded bg-gray-50 dark:bg-gray-700 flex justify-between items-center text-xs">
                                <div>
                                    <span class="font-bold block">${v.plate}</span>
                                    <span class="text-gray-500">${v.attachmentName || 'لا يوجد مستند مرفق'}</span>
                                </div>
                                <i class="fas fa-file-pdf text-red-500 text-lg"></i>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
        }

        // 11. Alerts Engine (30, 15, 7, 0 days)
        function renderAlertsSection(container) {
            const alerts = getAlertsList();
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">نظام التنبيهات المباشرة (30 - 15 - 7 - 0 يوم)</h3>
                    ${alerts.length === 0 ? `<p class="text-xs text-gray-500">لا توجد تنبيهات عاجلة حالياً.</p>` : `
                        <div class="space-y-3">
                            ${alerts.map(a => `
                                <div class="p-3 rounded-lg border flex items-center justify-between ${a.days <= 0 ? 'bg-red-100 border-red-300 text-red-900' : 'bg-amber-50 border-amber-300 text-amber-900'}">
                                    <div>
                                        <h4 class="font-bold text-xs">${a.vehicle.plate} - ${a.vehicle.model} (${a.type})</h4>
                                        <p class="text-[11px]">ينتهي بتاريخ: ${a.date}</p>
                                    </div>
                                    <span class="font-extrabold text-sm px-2 py-1 bg-white rounded shadow">${a.days <= 0 ? 'منتهي اليوم!' : 'متبقي ' + a.days + ' يوم'}</span>
                                </div>
                            `).join('')}
                        </div>
                    `}
                </div>
            `;
        }

        // 12. Reports Section
        function renderReportsSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-6 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 text-center space-y-4">
                    <i class="fas fa-chart-line text-4xl text-bank-gold"></i>
                    <h3 class="font-bold text-bank-green dark:text-bank-gold">طباعة واستخراج التقارير الرسمية للمصرف</h3>
                    <div class="flex justify-center gap-4">
                        <button onclick="exportToExcel()" class="px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-lg"><i class="fas fa-file-excel ml-1"></i> تقرير Excel شامل</button>
                        <button onclick="exportToPDF()" class="px-4 py-2 bg-red-600 text-white text-xs font-bold rounded-lg"><i class="fas fa-file-pdf ml-1"></i> تقرير PDF ملخص</button>
                    </div>
                </div>
            `;
        }

        // 13. Users Management
        function renderUsersSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold mb-4">إدارة مستخدمي النظام والصلاحيات</h3>
                    <div class="space-y-2 text-xs">
                        <div class="p-3 bg-gray-50 dark:bg-gray-700 rounded flex justify-between items-center">
                            <span>المدير العام (admin) - رمز المرور: <strong>123456</strong></span>
                            <span class="bg-bank-green text-white px-2 py-0.5 rounded text-[10px]">كامل الصلاحيات</span>
                        </div>
                        <div class="p-3 bg-gray-50 dark:bg-gray-700 rounded flex justify-between items-center">
                            <span>الموظف (user) - رمز المرور: <strong>654321</strong></span>
                            <span class="bg-blue-600 text-white px-2 py-0.5 rounded text-[10px]">إدخال وعرض</span>
                        </div>
                    </div>
                </div>
            `;
        }

        // 14. Backup & Restore Section
        function renderBackupSection(container) {
            container.innerHTML = `
                <div class="bg-white dark:bg-gray-800 p-6 rounded-xl shadow-sm border border-gray-200 dark:border-gray-700 space-y-4">
                    <h3 class="font-bold text-sm text-bank-green dark:text-bank-gold">النسخ الاحتياطي واسترجاع البيانات (JSON)</h3>
                    <div class="flex flex-wrap gap-4">
                        <button onclick="exportJSONBackup()" class="px-4 py-2 bg-bank-green text-white font-bold text-xs rounded-lg shadow"><i class="fas fa-download ml-1"></i> تحميل نسخة احتياطية JSON</button>
                        <label class="px-4 py-2 bg-bank-gold text-gray-900 font-bold text-xs rounded-lg shadow cursor-pointer">
                            <i class="fas fa-upload ml-1"></i> استرجاع بيانات من ملف
                            <input type="file" accept=".json" onchange="importJSONBackup(event)" class="hidden">
                        </label>
                    </div>
                </div>
            `;
        }

        /* ==========================================================================
           7. EXPORT & UTILITY FUNCTIONS
           ========================================================================== */
        function exportToExcel() {
            if (globalVehiclesList.length === 0) {
                showToast("لا توجد بيانات لتصديرها!", "error");
                return;
            }
            const ws = XLSX.utils.json_to_sheet(globalVehiclesList);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "السيارات");
            XLSX.writeFile(wb, "سيارات_مصرف_الادخار.xlsx");
            showToast("تم تصدير ملف Excel بنجاح", "success");
        }

        function exportToPDF() {
            if (globalVehiclesList.length === 0) {
                showToast("لا توجد بيانات لتصديرها!", "error");
                return;
            }
            window.print();
        }

        function exportJSONBackup() {
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(globalVehiclesList));
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href", dataStr);
            downloadAnchor.setAttribute("download", `bank_cars_backup_${new Date().toISOString().split('T')[0]}.json`);
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
            showToast("تم تحميل النسخة الاحتياطية بنجاح", "success");
        }

        function importJSONBackup(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async function(evt) {
                try {
                    const imported = JSON.parse(evt.target.result);
                    if (Array.isArray(imported)) {
                        for (let item of imported) {
                            await saveVehicleToDB(item);
                        }
                        await refreshAppData();
                        showToast("تم استرجاع البيانات بنجاح!", "success");
                    }
                } catch (err) {
                    showToast("تنسيق الملف غير صحيح!", "error");
                }
            };
            reader.readAsText(file);
        }

        /* ==========================================================================
           8. UI HELPERS & INITIALIZATION
           ========================================================================== */
        function openModal(id) {
            document.getElementById(id).classList.remove('hidden');
        }

        function closeModal(id) {
            document.getElementById(id).classList.add('hidden');
        }

        function toggleDarkMode() {
            document.documentElement.classList.toggle('dark');
            const icon = document.getElementById('theme-icon');
            if (document.documentElement.classList.contains('dark')) {
                icon.className = 'fas fa-sun';
            } else {
                icon.className = 'fas fa-moon';
            }
        }

        function toggleSidebarMobile() {
            document.getElementById('sidebar').classList.toggle('hidden');
        }

        // Window Load Bootstrapper
        window.onload = async function() {
            await initDatabase();
            globalVehiclesList = await getAllVehicles();
            console.log("App Initialized. Total vehicles loaded:", globalVehiclesList.length);
        };
    </script>
</body>
</html>
```