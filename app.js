import { initializeApp } from 'https://www.gstatic.com/firebasejs/9.22.1/firebase-app.js';
import { getFirestore, collection, addDoc, onSnapshot, query, where, serverTimestamp, updateDoc, doc, getDocs, orderBy, setDoc, getDocFromServer } from 'https://www.gstatic.com/firebasejs/9.22.1/firebase-firestore.js';
import { getAuth, signInAnonymously, onAuthStateChanged, signInWithEmailAndPassword } from 'https://www.gstatic.com/firebasejs/9.22.1/firebase-auth.js';

// ... (Service Data remains unchanged)

// Service Data
const SERVICES = [
    { id: 'anak', name: 'Anak', price: '35K', duration: 30 },
    { id: 'dewasa', name: 'Dewasa', price: '50K', duration: 45 },
    { id: 'semir', name: 'Semir Uban', price: '50K', duration: 30 },
    { id: 'downperm', name: 'Downperm', price: '120K', duration: 60 },
    { id: 'keratin', name: 'Keratin', price: '200K', duration: 90 },
    { id: 'perming', name: 'Perming Curly/Wavy', price: '250K', duration: 120 },
    { id: 'hairlight', name: 'Hairlight', price: '160K-200K', duration: 90 },
    { id: 'coloring', name: 'Coloring Full', price: '200K-250K', duration: 90 },
    { id: 'cornrows', name: 'Cornrows', price: '300K-500K', duration: 180 }
];

// Global State
let db, auth;
let selectedDate = new Date().toISOString().split('T')[0];
let selectedTime = null;
let occupiedSlotsByDay = {};
let allBookings = [];
let adminBookings = []; // Complete list for export

// DOM Elements
const servicesGrid = document.getElementById('services-grid');
const serviceSelect = document.getElementById('selected-service');
const slotsContainer = document.getElementById('slots-container');
const occupiedView = document.getElementById('occupied-slots-view');
const dateInput = document.getElementById('booking-date');
const bookingForm = document.getElementById('booking-form');
const adminFeed = document.getElementById('admin-feed-container');
const statRevenue = document.getElementById('foot-rev');
const statCustomers = document.getElementById('foot-cust');

// Helper for Firestore error handling as per system instructions
function handleFirestoreError(err, operationType, path = null) {
    if (err.code === 'permission-denied' || err.message?.includes('permission-denied')) {
        const user = auth.currentUser;
        const errorInfo = {
            error: String(err.message || 'Permission Denied'),
            operationType: String(operationType),
            path: path ? String(path) : null,
            authInfo: {
                userId: String(user ? user.uid : 'anonymous'),
                email: String(user ? user.email || '' : ''),
                emailVerified: Boolean(user ? user.emailVerified : false),
                isAnonymous: Boolean(user ? user.isAnonymous : true),
                providerInfo: user ? user.providerData.map(p => ({
                    providerId: String(p.providerId || ''),
                    displayName: String(p.displayName || ''),
                    email: String(p.email || '')
                })) : []
            }
        };
        
        let errorStr;
        try {
            errorStr = JSON.stringify(errorInfo);
        } catch (sErr) {
            errorStr = JSON.stringify({
                error: String(err.message || 'Circular Error Info'),
                operationType: String(operationType),
                path: String(path)
            });
        }
        
        console.error("Firestore Permission Error:", errorStr);
        throw errorStr;
    }
    console.error(`Firestore Error [${operationType}${path ? ': ' + path : ''}]:`, err.message || err);
    throw err;
}

async function startApp() {
    try {
        // Embedded config for zero-fetch initialization
        const firebaseConfig = {
            "projectId": "gen-lang-client-0398679268",
            "appId": "1:859454463198:web:4b37187bab2e62d9e7f64f",
            "apiKey": "AIzaSyAYUNyRQZPy12nrkUUV-Dj5n2t4S_znSV8",
            "authDomain": "gen-lang-client-0398679268.firebaseapp.com",
            "firestoreDatabaseId": "ai-studio-4d425c3f-fbba-4591-83d7-3e74f5a978fa",
            "storageBucket": "gen-lang-client-0398679268.firebasestorage.app",
            "messagingSenderId": "859454463198"
        };
        
        const app = initializeApp(firebaseConfig);
        db = getFirestore(app, firebaseConfig.firestoreDatabaseId);
        auth = getAuth(app);

        // Pre-flight Connection Check
        testFirebaseConnection();

        init();
        initShopStatus();
    } catch (err) {
        console.error("Failed to bootstrap application core:", err);
    }
}

async function testFirebaseConnection() {
    try {
        // Attempt to reach Cloud Firestore backend
        await getDocFromServer(doc(db, 'settings', 'store_status'));
        console.log("Firebase Backend: SYNCHRONIZED");
    } catch (error) {
        if (error.code === 'permission-denied' || error.message?.includes('permission-denied') || error.message?.includes('offline')) {
            console.error("CRITICAL: Firebase Connection Refused. Verify Domain Authorization.");
            console.warn("ADMIN ALERT: System isolation detected. Check Google Cloud Console Authorized Domains.");
        }
    }
}

let isShopOpen = true;
function initShopStatus() {
    // Sync with Firestore: settings/store_status (isOpen)
    onSnapshot(doc(db, 'settings', 'store_status'), (snapshot) => {
        const data = snapshot.data();
        isShopOpen = data ? data.isOpen : true;
        
        const display = document.getElementById('shop-status-display');
        const adminText = document.getElementById('shop-status-text');
        const adminBtn = document.getElementById('toggle-shop-btn');
        
        if (display) {
            display.classList.remove('hidden');
            display.className = `shop-status-banner ${isShopOpen ? 'shop-open' : 'shop-closed'}`;
            display.textContent = isShopOpen ? 'Sistem Aktif: Silahkan Booking' : 'Mohon Maaf, saat ini Skull Barber sedang tutup sementara. Silakan hubungi via WhatsApp untuk info lebih lanjut.';
        }

        if (adminText) {
            const statusLabel = isShopOpen ? 'BUKA' : 'TUTUP';
            adminText.textContent = statusLabel;
            adminText.className = isShopOpen ? 'text-green-500 font-black' : 'text-red-500 font-black';
        }

        if (adminBtn) {
            adminBtn.textContent = isShopOpen ? 'TUTUP TOKO (MODE ISTIRAHAT)' : 'BUKA TOKO SEKARANG';
            adminBtn.className = `px-8 py-3 text-[10px] font-black uppercase tracking-widest rounded transition-all shadow-lg ${isShopOpen ? 'bg-red-600 border border-red-700 text-white hover:bg-red-700' : 'bg-green-600 border border-green-700 text-white hover:bg-green-700'}`;
        }

        // Disable booking components if shop closed
        const bookingBtn = document.getElementById('submit-booking');
        const bookingInputs = document.querySelectorAll('#booking-form input, #booking-form select');
        
        if (bookingBtn) {
            bookingBtn.disabled = !isShopOpen;
            if (!isShopOpen) {
                bookingBtn.textContent = 'MODE ISTIRAHAT';
                bookingBtn.classList.add('opacity-30');
            } else {
                bookingBtn.textContent = 'Establish Appointment';
                bookingBtn.classList.remove('opacity-30');
            }
        }

        bookingInputs.forEach(input => {
            input.disabled = !isShopOpen;
            input.style.opacity = isShopOpen ? "1" : "0.4";
        });
    }, (err) => {
        handleFirestoreError(err, 'get', 'settings/store_status');
    });

    const toggleBtn = document.getElementById('toggle-shop-btn');
    if (toggleBtn) {
        toggleBtn.onclick = async () => {
            const originalText = toggleBtn.textContent;
            try {
                toggleBtn.disabled = true;
                toggleBtn.textContent = 'UPDATING STATE...';
                await setDoc(doc(db, 'settings', 'store_status'), { isOpen: !isShopOpen });
            } catch (err) {
                handleFirestoreError(err, 'write', 'settings/store_status');
            } finally {
                toggleBtn.disabled = false;
                toggleBtn.textContent = originalText;
            }
        };
    }
}

function init() {
    // Inject Services into Grid
    SERVICES.forEach(s => {
        const card = document.createElement('div');
        card.className = 'service-item-sleek';
        card.innerHTML = `
            <div class="service-header uppercase">
                <span class="service-name">${s.name}</span>
                <span class="service-price">${s.price}</span>
            </div>
            <div class="service-meta">${s.duration} Minutes • Professional Treatment</div>
        `;
        card.onclick = () => {
            serviceSelect.value = s.id;
            window.scrollToSection('main-layout');
        };
        servicesGrid.appendChild(card);

        const opt = document.createElement('option');
        opt.value = s.id;
        opt.textContent = `${s.name} - ${s.price}`;
        serviceSelect.appendChild(opt);
    });

    dateInput.value = selectedDate;
    dateInput.min = selectedDate;
    updateSlots();

    // Event Listeners
    dateInput.onchange = (e) => {
        selectedDate = e.target.value;
        updateSlots();
        listenToBookings();
    };

    bookingForm.onsubmit = handleBooking;
    
    // Admin Trigger (Internal)
    document.getElementById('login-btn').onclick = async () => {
        const email = document.getElementById('admin-email').value;
        const pass = document.getElementById('admin-pass').value;
        const btn = document.getElementById('login-btn');
        const err = document.getElementById('login-error');

        if (email !== 'skullstudio09@gmail.com') {
            err.textContent = 'RESTRICTED ACCESS ONLY';
            err.classList.remove('hidden');
            setTimeout(() => err.classList.add('hidden'), 3000);
            return;
        }

        try {
            btn.disabled = true;
            btn.textContent = 'VALIDATING...';
            await signInWithEmailAndPassword(auth, email, pass);
            document.getElementById('login-modal').classList.add('hidden');
            document.getElementById('admin-modal').classList.remove('hidden');
            initAdmin();
        } catch (authErr) {
            console.error(authErr);
            err.textContent = 'IDENTITY REJECTED';
            err.classList.remove('hidden');
            setTimeout(() => err.classList.add('hidden'), 3000);
        } finally {
            btn.disabled = false;
            btn.textContent = 'Authenticate';
        }
    };

    listenToBookings();
    signInAnonymously(auth).catch(err => {
        console.warn("Auth Provisioning Note:", err.message);
    });
}

function updateSlots() {
    slotsContainer.innerHTML = '';
    const day = new Date(selectedDate).getDay();
    const startHour = (day === 5) ? 13 : 10;
    const endHour = 22;

    for (let h = startHour; h < endHour; h++) {
        for (let m = 0; m < 60; m += 30) {
            const timeStr = `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
            const slot = document.createElement('div');
            const count = occupiedSlotsByDay[timeStr] || 0;
            const isFull = count >= 2;

            slot.className = `slot-sleek ${isFull ? 'full' : 'available'} ${selectedTime === timeStr ? 'selected' : ''}`;
            slot.innerHTML = `
                <span class="time">${timeStr}</span>
                <span class="status-label-sleek">${isFull ? 'Full' : (count + '/2 Slots')}</span>
            `;
            
            if (!isFull) {
                slot.onclick = () => {
                    selectedTime = timeStr;
                    updateSlots();
                };
            }
            slotsContainer.appendChild(slot);
        }
    }
}

let bookingsUnsubscribe = null;
function listenToBookings() {
    if (bookingsUnsubscribe) bookingsUnsubscribe();

    const q = query(collection(db, 'bookings'), where('date', '==', selectedDate));
    bookingsUnsubscribe = onSnapshot(q, (snapshot) => {
        occupiedSlotsByDay = {};
        allBookings = [];
        const publicFeed = document.getElementById('public-feed');
        if (publicFeed) publicFeed.innerHTML = '';

        snapshot.forEach(doc => {
            const data = doc.data();
            allBookings.push({ id: doc.id, ...data });

            if (data.status !== 'cancelled' && data.status !== 'no-show') {
                occupiedSlotsByDay[data.time] = (occupiedSlotsByDay[data.time] || 0) + 1;
                
                // Public Feed Injection
                if (publicFeed) {
                    const item = document.createElement('div');
                    item.className = 'ticket-mini-sleek';
                    item.innerHTML = `
                        <div class="time-stamp">${data.time} - BOOKED</div>
                        <div class="cust-name text-gold uppercase opacity-50">${data.customerName.charAt(0)}***${data.customerName.slice(-1)}</div>
                        <div class="service-type">${data.service} • Secured</div>
                    `;
                    publicFeed.appendChild(item);
                }
            }
        });
        updateSlots();
        updateLiveView();
    }, (err) => {
        handleFirestoreError(err, 'list', 'bookings');
    });
}

function updateLiveView() {
    occupiedView.innerHTML = '';
    Object.keys(occupiedSlotsByDay).sort().forEach(time => {
        const count = occupiedSlotsByDay[time];
        const dot = document.createElement('div');
        dot.className = `px-3 py-1 text-[9px] font-bold uppercase border ${count >= 2 ? 'bg-gold text-black border-gold' : 'border-zinc text-dim'}`;
        dot.textContent = `${time} ${count >= 2 ? 'FULL' : 'BUSY'}`;
        occupiedView.appendChild(dot);
    });
    if (Object.keys(occupiedSlotsByDay).length === 0) {
        occupiedView.innerHTML = '<p class="text-[9px] opacity-30 font-bold uppercase tracking-widest py-2">System Neutral: All Slots Available</p>';
    }
}

async function handleBooking(e) {
    e.preventDefault();

    const nameInput = document.getElementById('cust-name');
    const phoneInput = document.getElementById('cust-phone');
    const serviceId = document.getElementById('selected-service').value;

    // Validation Check
    if (!nameInput.value.trim()) {
        alert("Nama Pelanggan Wajib Diisi");
        nameInput.focus();
        return;
    }
    if (!phoneInput.value.trim() || phoneInput.value.length < 8) {
        alert("Nomor WhatsApp Tidak Valid");
        phoneInput.focus();
        return;
    }
    if (!selectedTime) {
        alert("Silahkan Pilih Jam Booking Terlebih Dahulu");
        window.scrollToSection('main-layout');
        return;
    }

    const service = SERVICES.find(s => s.id === serviceId);

    const bookingData = {
        customerName: nameInput.value.trim(),
        phoneNumber: phoneInput.value.trim(),
        service: service.name,
        serviceId: serviceId,
        price: service.price,
        date: selectedDate,
        time: selectedTime,
        status: 'pending',
        createdBy: auth.currentUser ? auth.currentUser.uid : 'anonymous',
        createdAt: serverTimestamp()
    };

    console.log("Attempting to send data to Firebase...");
    const btn = document.getElementById('submit-booking');

    try {
        btn.disabled = true;
        btn.textContent = 'EXECUTING TRANSACTION...';
        
        const docRef = await addDoc(collection(db, 'bookings'), bookingData);
        console.log("Transmission Success: ID", docRef.id);
        
        showTicket({ id: docRef.id, ...bookingData });
        bookingForm.reset();
        selectedTime = null;
        updateSlots();
    } catch (err) {
        console.error("Transmission Error:", err);
        // User-friendly feedback
        alert("Maaf, ada gangguan teknis. Hubungi admin via WhatsApp.");
        btn.textContent = 'Establish Appointment';
        btn.disabled = false;
        
        // Internal logging
        try {
            handleFirestoreError(err, 'create', 'bookings');
        } catch(e) { /* silent */ }
    } finally {
        if (!btn.disabled) {
            btn.textContent = 'Establish Appointment';
        }
    }
}

function showTicket(data) {
    document.getElementById('t-id').textContent = data.id || 'N/A';
    document.getElementById('t-date').textContent = data.date;
    document.getElementById('t-time').textContent = data.time;
    document.getElementById('t-name').textContent = data.customerName;
    document.getElementById('t-service').textContent = data.service;

    const waMsg = encodeURIComponent(`Halo Skull Barbershop, saya ingin konfirmasi booking:\n\nID: ${data.id}\nNama: ${data.customerName}\nLayanan: ${data.service}\nWaktu: ${data.date} jam ${data.time}\n\nTerima kasih!`);
    
    // Set both global and modal WA buttons
    const waGlobal = document.getElementById('wa-confirm-global');
    const waModal = document.getElementById('wa-confirm-modal');
    
    const openWA = () => window.open(`https://wa.me/6285723883091?text=${waMsg}`, '_blank');
    
    waGlobal.onclick = openWA;
    waModal.onclick = openWA;

    const cancelBtn = document.getElementById('cancel-booking-btn');
    cancelBtn.onclick = async () => {
        if (confirm("Are you sure you want to cancel this booking?")) {
            try {
                cancelBtn.disabled = true;
                cancelBtn.textContent = 'CANCELING...';
                await updateDoc(doc(db, 'bookings', data.id), { status: 'cancelled' });
                alert("Booking successfully cancelled.");
                document.getElementById('ticket-modal').classList.add('hidden');
            } catch (err) {
                handleFirestoreError(err, 'update', `bookings/${data.id}`);
            }
        }
    };

    document.getElementById('ticket-modal').classList.remove('hidden');
}

function initAdmin() {
    const q = query(collection(db, 'bookings'), orderBy('date', 'desc'), orderBy('time', 'desc'));
    onSnapshot(q, (snapshot) => {
        adminFeed.innerHTML = '';
        let revenue = 0;
        let customers = 0;
        let serviceCounts = {};
        
        const now = new Date();
        const currentTimeStr = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
        const todayStr = now.toISOString().split('T')[0];

        const all = [];
        snapshot.forEach(d => {
            const b = { id: d.id, ...d.data() };
            if (b.date === todayStr) {
                customers++;
                if (b.status === 'completed') {
                    revenue += parseInt(b.price.replace('K', '').split('-')[0]) * 1000;
                    serviceCounts[b.service] = (serviceCounts[b.service] || 0) + 1;
                }
            }
            all.push(b);
        });

        // Calculate Top Service
        let topService = '-';
        let maxCount = 0;
        for (const [srv, count] of Object.entries(serviceCounts)) {
            if (count > maxCount) {
                maxCount = count;
                topService = srv;
            }
        }

        all.forEach(booking => {
            const isToday = booking.date === todayStr;
            const hourToCheck = booking.time.split(':')[0];
            const minuteToCheck = booking.time.split(':')[1];
            const bookingTime = new Date();
            bookingTime.setHours(parseInt(hourToCheck), parseInt(minuteToCheck), 0);
            
            const diffInMs = bookingTime - now;
            const isActiveNow = isToday && Math.abs(diffInMs) <= (30 * 60 * 1000) && booking.status === 'pending'; // Active if within 30 mins

            const item = document.createElement('div');
            item.className = `ticket-mini-sleek ${isActiveNow ? 'active-glow' : (isToday && booking.status === 'pending' ? 'active-booking' : '')} flex flex-col md:flex-row justify-between items-start md:items-center gap-4 pr-4`;
            
            item.innerHTML = `
                <div class="w-full">
                    <div class="time-stamp badass-text text-[9px]">${booking.time} ${isActiveNow ? '- CURRENT SESSION' : (isToday ? '- TODAY' : '- ARCHIVED')}</div>
                    <div class="cust-name uppercase font-black text-white text-lg tracking-tight">${booking.customerName}</div>
                    <div class="service-type text-[10px] font-bold tracking-[1px] opacity-60 flex items-center gap-2">
                        <span>${booking.service.toUpperCase()}</span>
                        <span class="w-1 h-1 bg-zinc-600 rounded-full"></span>
                        <span class="font-mono text-gold">${booking.price}</span>
                    </div>
                </div>
                <div class="w-full md:w-auto flex flex-col sm:flex-row items-center gap-3 pt-4 md:pt-0">
                    <button onclick="window.openWA('${booking.phoneNumber}', '${booking.customerName}', '${booking.time}')" class="wa-admin-btn w-full sm:w-auto justify-center">
                        <svg class="w-3 h-3" fill="currentColor" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.67-1.611-.918-2.206-.242-.581-.487-.502-.67-.511-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413Z"/></svg>
                        WA
                    </button>
                    <select onchange="window.updateStatus('${booking.id}', this.value)" class="admin-action-select w-full sm:w-auto text-[10px] uppercase font-black">
                        <option value="pending" ${booking.status === 'pending' ? 'selected' : ''}>PENDING</option>
                        <option value="completed" ${booking.status === 'completed' ? 'selected' : ''}>SELESAI</option>
                        <option value="no-show" ${booking.status === 'no-show' ? 'selected' : ''}>NO-SHOW</option>
                        <option value="cancelled" ${booking.status === 'cancelled' ? 'selected' : ''}>CANCELLED</option>
                    </select>
                </div>
            `;
            adminFeed.appendChild(item);
        });

        // Update Dashboard Stats
        const dashRev = document.getElementById('dash-rev');
        const dashCust = document.getElementById('dash-cust');
        const dashTop = document.getElementById('dash-top-service');
        if (dashRev) dashRev.textContent = revenue.toLocaleString();
        if (dashCust) dashCust.textContent = customers;
        if (dashTop) dashTop.textContent = topService;

        statRevenue.textContent = revenue.toLocaleString();
        statCustomers.textContent = customers;
        adminBookings = all; // Update global export source
    }, (err) => {
        handleFirestoreError(err, 'list', 'bookings (admin)');
    });

    const exportBtn = document.getElementById('export-excel-btn');
    if (exportBtn) {
        exportBtn.onclick = handleExport;
    }
}

function handleExport() {
    if (!adminBookings || adminBookings.length === 0) {
        alert("No data available for export.");
        return;
    }

    const range = document.getElementById('export-range').value;
    const now = new Date();
    const todayStr = now.toISOString().split('T')[0];
    
    let filtered = adminBookings.filter(b => b.status === 'completed');
    
    if (range === 'today') {
        filtered = filtered.filter(b => b.date === todayStr);
    } else if (range === 'week') {
        const lastWeek = new Date();
        lastWeek.setDate(now.getDate() - 7);
        filtered = filtered.filter(b => new Date(b.date) >= lastWeek);
    } else if (range === 'month') {
        const currentMonth = now.getMonth();
        const currentYear = now.getFullYear();
        filtered = filtered.filter(b => {
            const d = new Date(b.date);
            return d.getMonth() === currentMonth && d.getFullYear() === currentYear;
        });
    }

    if (filtered.length === 0) {
        alert("No records found for the selected time range.");
        return;
    }

    // Transform data for SheetJS
    const reportData = filtered.map(b => ({
        'DATE': b.date,
        'TIME': b.time,
        'CLIENT': b.customerName,
        'PHONE (WA)': b.phoneNumber,
        'SERVICE': b.service,
        'PRICE': b.price,
        'STATUS': b.status.toUpperCase(),
        'BOOKED AT': b.createdAt ? new Date(b.createdAt.seconds * 1000).toLocaleString() : '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(reportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Recap");
    
    // Auto-size columns
    const max_width = reportData.reduce((w, r) => Math.max(w, Object.values(r).join('').length / 5), 10);
    worksheet["!cols"] = [ { wch: 15 }, { wch: 10 }, { wch: 20 }, { wch: 15 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 20 } ];

    XLSX.writeFile(workbook, `Skull_Recap_${range}_${todayStr}.xlsx`);
}

window.openWA = (phone, name, time) => {
    const msg = encodeURIComponent(`Halo ${name}, kami dari Skull Barbershop mengonfirmasi antrean Anda jam ${time}. Apakah Anda sudah di jalan?`);
    window.open(`https://wa.me/${phone.replace(/[^0-9]/g, '')}?text=${msg}`, '_blank');
};

window.updateStatus = async (id, status) => {
    try {
        await updateDoc(doc(db, 'bookings', id), { status });
    } catch (err) {
        handleFirestoreError(err, 'update', `bookings/${id}`);
    }
};

startApp();
