/* ============================================================
   Service Worker — Suivi financier
   ------------------------------------------------------------
   SW_VERSION DOIT être copiée depuis APP_VERSION (index.html) à
   CHAQUE livraison. C'est ce qui force le navigateur à détecter
   que le fichier a changé (comparaison octet à octet du script)
   et donc à déclencher le cycle install/activate qui invalide
   l'ancien cache — sans ça, une page déjà ouverte pourrait
   rester bloquée sur une version obsolète indéfiniment.
   ============================================================ */
const SW_VERSION = "v 2026.09.26.09.22";

const SHELL_CACHE   = `sf-shell-${SW_VERSION}`;
const RUNTIME_CACHE = "sf-runtime"; // ressources externes (CDN) — non versionné, survit aux mises à jour

const SHELL_FILES = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-maskable-192.png",
  "./icon-maskable-512.png",
  "./apple-touch-icon.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE)
      .then((cache) => cache.addAll(SHELL_FILES))
      .catch(() => {}) // ne bloque pas l'installation si une ressource est indisponible au premier essai
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((names) => Promise.all(
        names
          .filter((name) => name.startsWith("sf-shell-") && name !== SHELL_CACHE)
          .map((name) => caches.delete(name))
      ))
      .then(() => self.clients.claim())
  );
});

// La page (index.html) envoie ce message quand l'utilisateur confirme
// la mise à jour depuis la bannière "Nouvelle version disponible".
self.addEventListener("message", (event) => {
  if (event.data === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  const isSameOrigin = url.origin === location.origin;

  // Page principale : réseau d'abord (toujours la dernière version si en ligne),
  // repli sur le cache si hors ligne.
  if (req.mode === "navigate" || (isSameOrigin && url.pathname.endsWith("index.html"))) {
    event.respondWith(
      fetch(req)
        .then((res) => {
          // Ne met en cache que les réponses same-origin valides (jamais une page
          // d'erreur GitHub Pages qui remplacerait silencieusement le shell en cache).
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(SHELL_CACHE).then((cache) => cache.put(req, copy));
          }
          return res;
        })
        .catch(() => caches.match(req).then((res) => res || caches.match("./index.html")))
    );
    return;
  }

  // Fichiers de l'app shell (manifest, icônes) : cache d'abord, réseau en secours.
  if (isSameOrigin) {
    event.respondWith(
      caches.match(req).then((cached) => cached || fetch(req))
    );
    return;
  }

  // Ressources externes (Google Fonts, SheetJS/cdnjs) : stale-while-revalidate
  // pour permettre un usage hors ligne après un premier chargement en ligne.
  event.respondWith(
    caches.open(RUNTIME_CACHE).then((cache) =>
      cache.match(req).then((cached) => {
        const fetchPromise = fetch(req)
          .then((res) => {
            // Ne met en cache que les réponses exploitables (200 ou opaque cross-origin) :
            // jamais une erreur 4xx/5xx du CDN, qui remplacerait durablement une copie valide.
            if (res && (res.ok || res.type === "opaque")) cache.put(req, res.clone());
            return res;
          })
          .catch(() => cached);
        return cached || fetchPromise;
      })
    )
  );
});
