const CACHE_NAME = 'presure-cache-v2.1';

const urlsToCache = [
  './',
  './index.html',
  './styles.css', // ¡Añadido: Estilos!
  './app.js',     // ¡Añadido: Lógica!
  './manifest.json',
  './icon-192.png',
  './apple-touch-icon.png',
  './favicon-32x32.png',
  './icon-512.png',
  './icon-maskable-512.png',

  // Librerías externas (CDNs)
  'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.29/jspdf.plugin.autotable.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/localforage/1.10.0/localforage.min.js',

  // Font Awesome
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.woff2',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/webfonts/fa-solid-900.ttf'
];

// 1. INSTALACIÓN
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('Service Worker: Cacheando archivos de PresuRE');
        return cache.addAll(urlsToCache);
      })
  );
});

// 2. ACTIVACIÓN (Limpia cachés antiguas)
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cache => {
          if (cache !== CACHE_NAME) {
            console.log('Service Worker: Limpiando caché antigua', cache);
            return caches.delete(cache);
          }
        })
      );
    })
  );
});

// 3. FETCH (Intercepta peticiones y sirve la caché si estás offline)
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Devuelve el archivo desde la caché si existe; si no, intenta ir a internet
        return response || fetch(event.request);
      })
      .catch(error => {
        console.error('Fetch failed, offline and no cache match:', error);
        // Opcional: Podrías devolver un fallback HTML aquí si falla la red y no está en caché
      })
  );
});