/* MandaGroup PWA — Initialization (dipisah dari app.js) */
(function () {
  'use strict';

  // --- Service Worker registration ---
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker
        .register('./sw.js', { scope: './' })
        .then(reg => {
          console.log('[PWA] SW terdaftar:', reg.scope);

          // Deteksi update
          reg.addEventListener('updatefound', () => {
            const sw = reg.installing;
            if (!sw) return;
            sw.addEventListener('statechange', () => {
              if (sw.state === 'installed' && navigator.serviceWorker.controller) {
                // Versi baru siap — tanya user
                if (window.confirm('Versi baru MandaGroup tersedia. Muat ulang sekarang?')) {
                  sw.postMessage('SKIP_WAITING');
                  window.location.reload();
                }
              }
            });
          });
        })
        .catch(err => console.warn('[PWA] Registrasi SW gagal:', err));
    });
  }

  // --- Install prompt (Android Chrome / Edge / Desktop) ---
  let deferredPrompt = null;

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    const btn = document.getElementById('pwa-install');
    if (btn) btn.hidden = false;
  });

  window.addEventListener('appinstalled', () => {
    console.log('[PWA] Aplikasi terpasang');
    deferredPrompt = null;
    const btn = document.getElementById('pwa-install');
    if (btn) btn.hidden = true;
  });

  // Tombol install manual (opsional, bisa ditaruh di top bar)
  window.installPWA = async function () {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    console.log('[PWA] Install:', outcome);
    deferredPrompt = null;
    const btn = document.getElementById('pwa-install');
    if (btn) btn.hidden = true;
  };
})();