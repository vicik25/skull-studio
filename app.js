import { initializeApp } from 'https://www.gstatic.com/firebasejs/9.22.1/firebase-app.js';
import { getFirestore, collection, addDoc, onSnapshot, query, where, serverTimestamp, updateDoc, doc, getDocs, orderBy, setDoc, getDocFromServer } from 'https://www.gstatic.com/firebasejs/9.22.1/firebase-firestore.js';
import { getAuth, signInAnonymously, onAuthStateChanged, signInWithEmailAndPassword, setPersistence, browserLocalPersistence } from 'https://www.gstatic.com/firebasejs/9.22.1/firebase-auth.js';

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
let todayStr = new Date().toISOString().split('T')[0];
let selectedTime = null;
let occupiedSlotsByDay = {};
let allBookings = [];
let adminBookings = []; // Complete list for export
let currentFeedbackId = null;
let selectedRating = 0;

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
const feedbackModal = document.getElementById('feedback-modal');
const starBtns = document.querySelectorAll('.star-btn');
const submitFeedbackBtn = document.getElementById('submit-feedback');
const feedbackText = document.getElementById('feedback-text');

// Helper for Firestore error handling
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

        // Crucial for keeping session on refresh
        await setPersistence(auth, browserLocalPersistence);

        testFirebaseConnection();

        init();
        initShopStatus();
        listenToTodayStats(); // Always listen to today's stats for the floating footer

        onAuthStateChanged(auth, (user) => {
            if (user && user.email === 'skullstudio09@gmail.com') {
                console.log("Admin session restored.");
                document.getElementById('login-modal').classList.add('hidden');
                document.getElementById('admin-modal').classList.remove('hidden');
                initAdmin();
            } else if (!user) {
                // Only sign in anonymously if we are truly logged out
                signInAnonymously(auth).catch(err => {
                    console.warn("Guest session entry failed:", err);
                });
            }
        });

    } catch (err) {
        console.error("Failed to bootstrap application core:", err);
    }
}

async function testFirebaseConnection() {
    try {
        await getDocFromServer(doc(db, 'settings', 'store_status'));
    } catch (error) {
        console.warn("Connection test concluded.");
    }
}

let isShopOpen = true;
function initShopStatus() {
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
            adminText.textContent = isShopOpen ? 'BUKA' : 'TUTUP';
            adminText.className = isShopOpen ? 'text-green-500 font-black' : 'text-red-500 font-black';
        }

        if (adminBtn) {
            adminBtn.textContent = isShopOpen ? 'TUTUP TOKO (MODE ISTIRAHAT)' : 'BUKA TOKO SEKARANG';
            adminBtn.className = `px-8 py-3 text-[10px] font-black uppercase tracking-widest rounded transition-all shadow-lg ${isShopOpen ? 'bg-red-600 border border-red-700 text-white hover:bg-red-700' : 'bg-green-600 border border-green-700 text-white hover:bg-green-700'}`;
        }

        const bookingBtn = document.getElementById('submit-booking');
        const bookingInputs = document.querySelectorAll('#booking-form input, #booking-form select');
        
        if (bookingBtn) {
            bookingBtn.disabled = !isShopOpen;
            bookingBtn.textContent = isShopOpen ? 'Establish Appointment' : 'MODE ISTIRAHAT';
            bookingBtn.classList.toggle('opacity-30', !isShopOpen);
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
            try {
                toggleBtn.disabled = true;
                await setDoc(doc(db, 'settings', 'store_status'), { isOpen: !isShopOpen });
            } catch (err) {
                handleFirestoreError(err, 'write', 'settings/store_status');
            } finally {
                toggleBtn.disabled = false;
            }
        };
    }
}

function init() {
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

    dateInput.onchange = (e) => {
        selectedDate = e.target.value;
        updateSlots();
        listenToBookings();
    };

    bookingForm.onsubmit = handleBooking;
    
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
            err.textContent = 'IDENTITY REJECTED';
            err.classList.remove('hidden');
            setTimeout(() => err.classList.add('hidden'), 3000);
        } finally {
            btn.disabled = false;
            btn.textContent = 'Authenticate';
        }
    };

    starBtns.forEach(btn => {
        btn.onclick = () => {
            selectedRating = parseInt(btn.dataset.rate);
            starBtns.forEach((b, idx) => {
                b.classList.toggle('text-gold', idx < selectedRating);
                b.classList.toggle('text-zinc-700', idx >= selectedRating);
            });
        };
    });

    submitFeedbackBtn.onclick = handleFeedbackSubmit;

    listenToBookings();
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
            const isFull = count >= 4;

            slot.className = `slot-sleek ${isFull ? 'full' : 'available'} ${selectedTime === timeStr ? 'selected' : ''}`;
            slot.innerHTML = `
                <span class="time">${timeStr}</span>
                <span class="status-label-sleek">${isFull ? 'Full' : (count + '/4 Slots')}</span>
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
let adminUnsubscribe = null;
let statsUnsubscribe = null;

// New dedicated listener for Today's Stats (Revenue & Customer Count)
function listenToTodayStats() {
    if (statsUnsubscribe) statsUnsubscribe();

    const q = query(collection(db, 'bookings'), where('date', '==', todayStr));
    statsUnsubscribe = onSnapshot(q, (snapshot) => {
        let revenue = 0;
        let customers = 0;

        snapshot.forEach(d => {
            const data = d.data();
            // Count for today's active customers (non-cancelled/no-show)
            if (data.status !== 'cancelled' && data.status !== 'no-show') {
                customers++;
                // Revenue only from completed ones
                if (data.status === 'completed') {
                    revenue += parseInt(data.price.replace('K', '').split('-')[0]) * 1000;
                }
            }
        });

        // Update Floating Footer Stats
        if (statRevenue) statRevenue.textContent = revenue.toLocaleString();
        if (statCustomers) statCustomers.textContent = customers;

        // Also update Main Dashboard if we are in admin view
        const dashRev = document.getElementById('dash-rev');
        const dashCust = document.getElementById('dash-cust');
        if (dashRev) dashRev.textContent = revenue.toLocaleString();
        if (dashCust) dashCust.textContent = customers;
    }, (err) => {
        console.error("Stats Listener Error:", err);
    });
}

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
            
            // Occupy slot if not cancelled/no-show
            if (data.status === 'pending' || data.status === 'completed') {
                occupiedSlotsByDay[data.time] = (occupiedSlotsByDay[data.time] || 0) + 1;
            }
            
            if (publicFeed && data.status !== 'cancelled' && data.status !== 'no-show') {
                const item = document.createElement('div');
                item.className = 'ticket-mini-sleek';
                item.innerHTML = `
                    <div class="time-stamp">${data.time} - ${data.status.toUpperCase()}</div>
                    <div class="cust-name text-gold uppercase opacity-50">${data.customerName.charAt(0)}***${data.customerName.slice(-1)}</div>
                    <div class="service-type">${data.service} • Secured</div>
                `;
                publicFeed.appendChild(item);
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
        dot.className = `px-3 py-1 text-[9px] font-bold uppercase border ${count >= 4 ? 'bg-gold text-black border-gold' : 'border-zinc text-dim'}`;
        dot.textContent = `${time} ${count >= 4 ? 'FULL' : 'BUSY'}`;
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

    const barberSelect = document.getElementById('selected-barber');

    if (!nameInput.value.trim()) { alert("Nama Pelanggan Wajib Diisi"); nameInput.focus(); return; }
    if (!phoneInput.value.trim() || phoneInput.value.length < 8) { alert("Nomor WhatsApp Tidak Valid"); phoneInput.focus(); return; }
    if (!selectedTime) { alert("Silahkan Pilih Jam Booking Terlebih Dahulu"); return; }

    const service = SERVICES.find(s => s.id === serviceId);

    if (!auth.currentUser) {
        try { await signInAnonymously(auth); } catch (authErr) { alert("Sesi gagal dimulai."); return; }
    }

    const bookingData = {
        customerName: nameInput.value.trim(),
        phoneNumber: phoneInput.value.trim(),
        service: service.name,
        serviceId: serviceId,
        barberName: barberSelect.value, // Added barber selection
        price: service.price,
        date: selectedDate,
        time: selectedTime,
        status: 'pending',
        createdBy: auth.currentUser.uid,
        createdAt: serverTimestamp()
    };

    const btn = document.getElementById('submit-booking');
    try {
        btn.disabled = true;
        btn.textContent = 'EXECUTING TRANSACTION...';
        const docRef = await addDoc(collection(db, 'bookings'), bookingData);
        showTicket({ id: docRef.id, ...bookingData });
        bookingForm.reset();
        selectedTime = null;
        updateSlots();
    } catch (err) {
        handleFirestoreError(err, 'create', 'bookings');
    } finally {
        btn.disabled = false;
        btn.textContent = 'Establish Appointment';
    }
}

function showTicket(data) {
    document.getElementById('t-id').textContent = data.id || 'N/A';
    document.getElementById('t-date').textContent = data.date;
    document.getElementById('t-time').textContent = data.time;
    document.getElementById('t-name').textContent = data.customerName;
    document.getElementById('t-service').textContent = data.service;
    
    // Display barber name if exists
    const tBarber = document.getElementById('t-barber');
    if (tBarber) tBarber.textContent = data.barberName || 'Siapa Saja';
    else {
        // If element doesn't exist in HTML yet, we could add it dynamically or just skip
        // But requested to show it, so let's check if we should add it to HTML too
    }

    const waMsg = encodeURIComponent(`Halo Skull Barbershop, saya ingin konfirmasi booking:\n\nID: ${data.id}\nNama: ${data.customerName}\nLayanan: ${data.service}\nBarber: ${data.barberName || 'Siapa Saja'}\nWaktu: ${data.date} jam ${data.time}\n\nTerima kasih!`);
    const openWA = () => window.open(`https://wa.me/6282134504657?text=${waMsg}`, '_blank');
    document.getElementById('wa-confirm-global').onclick = openWA;
    document.getElementById('wa-confirm-modal').onclick = openWA;

    const cancelBtn = document.getElementById('cancel-booking-btn');
    
    const unsub = onSnapshot(doc(db, 'bookings', data.id), (snapshot) => {
        const d = snapshot.data();
        if (d && (d.status === 'completed' || d.status === 'cancelled') && !d.feedback) {
            unsub(); 
            document.getElementById('ticket-modal').classList.add('hidden');
            openFeedback(data.id);
        }
    });

    cancelBtn.onclick = async () => {
        if (confirm("KONFIRMASI PEMBATALAN: Apakah Anda yakin ingin membatalkan antrean ini?")) {
            try {
                cancelBtn.disabled = true;
                await updateDoc(doc(db, 'bookings', data.id), { status: 'cancelled' });
                document.getElementById('ticket-modal').classList.add('hidden');
            } catch (err) {
                handleFirestoreError(err, 'update', `bookings/${data.id}`);
            } finally {
                cancelBtn.disabled = false;
            }
        }
    };

    document.getElementById('ticket-modal').classList.remove('hidden');
}

async function handleFeedbackSubmit() {
    if (!currentFeedbackId || selectedRating === 0) {
        alert("Silahkan berikan rating bintang terlebih dahulu.");
        return;
    }
    try {
        submitFeedbackBtn.disabled = true;
        await updateDoc(doc(db, 'bookings', currentFeedbackId), {
            feedback: `${selectedRating} Bintang - ${feedbackText.value.trim()}`
        });
        alert("Terima kasih atas feedback Anda!");
        closeModal('feedback-modal');
    } catch (err) {
        alert("Gagal mengirim feedback.");
    } finally {
        submitFeedbackBtn.disabled = false;
    }
}

function openFeedback(bookingId) {
    currentFeedbackId = bookingId;
    selectedRating = 0;
    feedbackText.value = '';
    starBtns.forEach(b => { b.classList.remove('text-gold'); b.classList.add('text-zinc-700'); });
    feedbackModal.classList.remove('hidden');
}

function initAdmin() {
    if (adminUnsubscribe) adminUnsubscribe();

    // Admins need to see all recent bookings for management and recap
    const q = query(collection(db, 'bookings'), orderBy('date', 'desc'), orderBy('time', 'desc'));
    adminUnsubscribe = onSnapshot(q, (snapshot) => {
        adminFeed.innerHTML = '';
        let serviceCounts = {};
        const now = new Date();
        const all = [];

        snapshot.forEach(d => {
            const b = { id: d.id, ...d.data() };
            if (b.date === todayStr && b.status === 'completed') {
                serviceCounts[b.service] = (serviceCounts[b.service] || 0) + 1;
            }
            all.push(b);
        });

        let topService = '-';
        let maxCount = 0;
        for (const [srv, count] of Object.entries(serviceCounts)) {
            if (count > maxCount) { maxCount = count; topService = srv; }
        }

        all.sort((a, b) => {
            // Priority 1: Pending first
            const aIsActive = a.status === 'pending';
            const bIsActive = b.status === 'pending';
            if (aIsActive && !bIsActive) return -1;
            if (!aIsActive && bIsActive) return 1;
            
            // Priority 2: Within same status, sort by date/time
            const aTimeStr = `${a.date}T${a.time}:00`;
            const bTimeStr = `${b.date}T${b.time}:00`;
            return aIsActive ? aTimeStr.localeCompare(bTimeStr) : bTimeStr.localeCompare(aTimeStr);
        });

        all.forEach(booking => {
            const isToday = booking.date === todayStr;
            const bookingDateTime = new Date(`${booking.date}T${booking.time}:00`);
            const expireThreshold = new Date(bookingDateTime.getTime() + (60 * 60 * 1000));
            const isExpired = booking.status === 'pending' && now > expireThreshold;
            
            const diffInMs = bookingDateTime - now;
            const isActiveNow = isToday && Math.abs(diffInMs) <= (30 * 60 * 1000) && booking.status === 'pending';

            const item = document.createElement('div');
            
            // Status-based styling classes
            let statusClasses = 'border-zinc-800 bg-zinc-900/30';
            if (booking.status === 'pending') {
                statusClasses = isExpired 
                    ? 'border-red-500/50 bg-red-900/10 shadow-[0_0_15px_rgba(239,68,68,0.1)]' 
                    : 'border-gold/30 bg-gold/5';
                if (isActiveNow) statusClasses += ' active-glow border-gold';
            } else if (booking.status === 'completed') {
                statusClasses = 'opacity-40 grayscale';
            } else if (booking.status === 'no-show') {
                statusClasses = 'border-red-900/50 bg-red-950/20 opacity-60';
            }

            item.className = `ticket-mini-sleek border p-4 mb-3 transition-all flex flex-col md:flex-row justify-between items-start md:items-center gap-4 ${statusClasses}`;
            
            const statusLabel = isExpired ? '<span class="text-red-500 animate-pulse">WAKTU HABIS / PERLU KONFIRMASI</span>' : booking.status.toUpperCase();

            item.innerHTML = `
                <div class="w-full">
                    <div class="time-stamp badass-text text-[9px] mb-1 font-black tracking-widest ${isExpired ? 'text-red-500' : 'text-gold'}">
                        ${booking.time} • ${statusLabel} ${isToday ? '• TODAY' : ''}
                    </div>
                    <div class="cust-name uppercase font-black text-white text-lg tracking-tight flex items-center gap-2">
                        ${booking.customerName}
                        ${isActiveNow ? '<span class="w-2 h-2 bg-gold rounded-full animate-ping"></span>' : ''}
                    </div>
                    <div class="service-type text-[10px] font-bold tracking-[1px] opacity-60 flex flex-wrap items-center gap-2">
                        <span>${booking.service.toUpperCase()}</span>
                        <span class="w-1 h-1 bg-zinc-600 rounded-full"></span>
                        <span class="text-gold">BARBER: ${booking.barberName ? booking.barberName.toUpperCase() : 'SIAPA SAJA'}</span>
                        <span class="w-1 h-1 bg-zinc-600 rounded-full"></span>
                        <span class="font-mono text-gold">${booking.price}</span>
                        ${booking.feedback ? `<span class="ml-2 text-gold">★ ${booking.feedback}</span>` : ''}
                    </div>
                </div>
                <div class="w-full md:w-auto flex flex-wrap items-center gap-2 pt-4 md:pt-0">
                    <button onclick="window.openWA('${booking.phoneNumber}', '${booking.customerName}', '${booking.time}')" 
                        class="px-4 py-2 bg-green-600/20 border border-green-600/40 text-green-500 text-[10px] font-black uppercase tracking-widest hover:bg-green-600 hover:text-white transition-all rounded">
                        WA
                    </button>
                    
                    ${booking.status === 'pending' ? `
                        <button onclick="window.updateStatus('${booking.id}', 'completed')" 
                            class="px-4 py-2 bg-gold text-black text-[10px] font-black uppercase tracking-widest hover:scale-105 transition-all rounded shadow-lg shadow-gold/20">
                            Selesai
                        </button>
                        <button onclick="window.updateStatus('${booking.id}', 'no-show')" 
                            class="px-4 py-2 bg-red-600/20 border border-red-600/40 text-red-500 text-[10px] font-black uppercase tracking-widest hover:bg-red-600 hover:text-white transition-all rounded">
                            Tidak Datang
                        </button>
                    ` : `
                        <div class="relative">
                            <select onchange="window.updateStatus('${booking.id}', this.value)" class="admin-action-select text-[10px] uppercase font-black bg-transparent border border-zinc-700 p-2 rounded">
                                <option value="pending" ${booking.status === 'pending' ? 'selected' : ''}>PENDING</option>
                                <option value="completed" ${booking.status === 'completed' ? 'selected' : ''}>SELESAI</option>
                                <option value="no-show" ${booking.status === 'no-show' ? 'selected' : ''}>NO-SHOW</option>
                                <option value="cancelled" ${booking.status === 'cancelled' ? 'selected' : ''}>CANCELLED</option>
                            </select>
                        </div>
                    `}
                </div>
            `;
            adminFeed.appendChild(item);
        });

        const dashTopSrv = document.getElementById('dash-top-service');
        if (dashTopSrv) dashTopSrv.textContent = topService;
        adminBookings = all;
    }, (err) => {
        handleFirestoreError(err, 'list', 'bookings (admin)');
    });

    document.getElementById('export-excel-btn').onclick = handleExport;
}

function handleExport() {
    if (!adminBookings || adminBookings.length === 0) { alert("No data available."); return; }
    const range = document.getElementById('export-range').value;
    const now = new Date();
    const todayStr = now.toISOString().split('T')[0];
    let filtered = [...adminBookings];
    if (range === 'today') filtered = filtered.filter(b => b.date === todayStr);
    else if (range === 'week') { const lastWeek = new Date(); lastWeek.setDate(now.getDate() - 7); filtered = filtered.filter(b => new Date(b.date) >= lastWeek); }
    else if (range === 'month') { const currentMonth = now.getMonth(); const currentYear = now.getFullYear(); filtered = filtered.filter(b => { const d = new Date(b.date); return d.getMonth() === currentMonth && d.getFullYear() === currentYear; }); }
    if (filtered.length === 0) { alert("No records found."); return; }

    const reportData = filtered.map(b => ({
        'DATE': b.date,
        'TIME': b.time,
        'CLIENT': b.customerName,
        'PHONE (WA)': b.phoneNumber,
        'SERVICE': b.service,
        'PRICE': b.price,
        'STATUS': b.status.toUpperCase(),
        'FEEDBACK': b.feedback || '-',
        'BOOKED AT': b.createdAt ? new Date(b.createdAt.seconds * 1000).toLocaleString() : '-'
    }));

    const worksheet = XLSX.utils.json_to_sheet(reportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Recap");
    worksheet["!cols"] = [ { wch: 15 }, { wch: 10 }, { wch: 20 }, { wch: 15 }, { wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 30 }, { wch: 20 } ];
    XLSX.writeFile(workbook, `Skull_Recap_${range}_${todayStr}.xlsx`);
}

function formatPhone(num) {
    if (!num) return '';
    let cleaned = num.replace(/\D/g, ''); // Hapus semua karakter non-digit
    if (cleaned.startsWith('0')) {
        cleaned = '62' + cleaned.slice(1); // Ubah 08xxx menjadi 628xxx
    }
    return cleaned;
}

window.openWA = (phone, name, time) => {
    const formattedPhone = formatPhone(phone);
    const msg = encodeURIComponent(`Halo Kak ${name}, ini dari Skull Barber Studio. Kami ingin mengonfirmasi antrean Anda untuk jam ${time}. Apakah ada yang bisa kami bantu?`);
    window.open(`https://wa.me/${formattedPhone}?text=${msg}`, '_blank');
};

window.updateStatus = async (id, status) => {
    if (status === 'cancelled' && !confirm("PERINGATAN ADMIN: Anda akan membatalkan pesanan ini?")) return;
    try { await updateDoc(doc(db, 'bookings', id), { status }); } catch (err) { handleFirestoreError(err, 'update', `bookings/${id}`); }
};

window.closeModal = (id) => { document.getElementById(id).classList.add('hidden'); };

startApp();
