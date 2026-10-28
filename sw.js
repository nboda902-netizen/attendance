// ضروري عشان المتصفح يعتبر الصفحة تطبيق قابل للتثبيت. مفيش تخزين مؤقت: التسجيل محتاج إنترنت دايماً.
self.addEventListener('install',function(){self.skipWaiting()});
self.addEventListener('activate',function(e){e.waitUntil(self.clients.claim())});
self.addEventListener('fetch',function(){});
