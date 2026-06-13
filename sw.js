const CACHE_NAME = 'converter-v1';
const ASSETS = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  './icône-192.png',
  './icône-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css',
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600&display=swap'
];

// Événement 'install' : Mise en cache des ressources essentielles
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('Service Worker : Mise en cache des ressources...');
      return cache.addAll(ASSETS);
    }).then(() => self.skipWaiting()) // Force le SW à devenir actif immédiatement
  );
});

// Événement 'activate' : Nettoyage des anciens caches pour éviter les conflits de version
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            console.log('Service Worker : Suppression de l\'ancien cache', cache);
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim()) // Prend le contrôle des pages ouvertes immédiatement
  );
});

// Événement 'fetch' : Stratégie Cache-First avec repli sur le réseau
self.addEventListener('fetch', (event) => {
  // Optionnel : Ignorer les requêtes non GET (comme les extensions de navigateur ou requêtes POST)
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      // Si la ressource est dans le cache, on la retourne, sinon on va la chercher sur le réseau
      return cachedResponse || fetch(event.request).catch(() => {
        // Optionnel : Ici tu pourrais retourner une page d'erreur personnalisée si le réseau échoue
        console.log('Ressource non trouvée dans le cache et réseau indisponible.');
      });
    })
  );
});
