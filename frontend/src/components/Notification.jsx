export default function Notification({ notification }) {
  if (!notification) return null;
  return (
    <div className={`fixed top-5 right-5 z-50 px-6 py-3 rounded-lg shadow-xl font-medium transition-all transform duration-300 translate-y-0 ${
      notification.type === 'error' 
        ? 'bg-rose-500/90 text-white border border-rose-400' 
        : 'bg-emerald-500/90 text-white border border-emerald-400'
    }`}>
      <div className="flex items-center gap-2">
        <span>{notification.type === 'error' ? '⚠️' : '✅'}</span>
        <span>{notification.message}</span>
      </div>
    </div>
  );
}
