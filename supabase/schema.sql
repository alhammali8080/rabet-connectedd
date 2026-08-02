-- قاعدة البيانات الدنيا لحفظ حالة رابط.
-- الخادم يستخدم service_role key، لذلك لا تضع هذا المفتاح في الواجهة أبداً.
create table if not exists public.app_state (
  id integer primary key check (id = 1),
  state jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.app_state enable row level security;

-- لا نفتح القراءة العامة. الخادم فقط يقرأ ويكتب عبر service role.
-- لا تُنشئ policies عامة لهذا الجدول في نسخة الإنتاج.
