import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  LayoutDashboard, 
  Package, 
  Store as StoreIcon, 
  Users, 
  FileText, 
  LogOut, 
  Plus, 
  Search, 
  MapPin, 
  Check, 
  TrendingUp, 
  DollarSign, 
  ShoppingCart, 
  UserCheck,
  AlertCircle,
  ChevronDown,
  Settings,
  Sun,
  Moon,
  Edit,
  Trash2,
  Menu,
  X,
  Camera,
  CheckCircle,
  CreditCard,
  Loader,
  Clock
} from 'lucide-react';
import { Html5Qrcode } from 'html5-qrcode';
import { 
  AreaChart, 
  Area, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  BarChart,
  Bar
} from 'recharts';

// Mock Initial Data for fully interactive offline testing

const translations = {
  uz: {
    login_title: "Tizimga kirish",
    login_subtitle: "Ombor va agentlarni boshqarish paneli",
    username: "Foydalanuvchi nomi",
    password: "Parol",
    login_btn: "Kirish",
    login_error: "Foydalanuvchi nomi yoki parol xato!",
    demo_mode: "Demo: admin / admin (To'liq admin) yoki ombor / ombor (Omborchi)",
    dashboard: "Asosiy Panel",
    products: "Mahsulotlar",
    stores: "Do'konlar",
    agents: "Agentlar boshqaruvi",
    analytics: "Tahlillar",
    general_analytics: "Umumiy tahlil",
    store_analytics: "Do'kon tahlili",
    agent_analytics: "Agent tahlillari",
    assignments: "Agentlarga biriktirish",
    my_tasks: "Mening vazifalarim",
    sales_history: "Sotuvlar tarixi",
    settings: "Sozlamalar",
    profile_settings: "Profil sozlamalari",
    language_settings: "Til sozlamalari",
    logout: "Chiqish",
    active_in_system: "Tizimda faol",
    date_filter: "Tahlil muddati",
    bugun: "Bugun",
    dan: "dan",
    gacha: "gacha",
    custom_range: "Muddat tanlash",
    kechagi: "Kechagi",
    "3-kunlik": "3 kunlik",
    "7-kunlik": "7 kunlik",
    "15-kunlik": "15 kunlik",
    "30-kunlik": "30 kunlik",
    hammasi: "Hammasi",
    revenue: "Savdo tushumi",
    active_agents: "Faol agentlar",
    sold_items: "Sotilgan tovarlar",
    net_profit: "Sof foyda",
    weekly_dynamics: "Haftalik sotuvlar dinamikasi",
    active_agents_today: "Bugungi faol agentlar",
    agent_name: "Ismi",
    stores_count: "Do'konlar soni",
    total_sales: "Jami sotuv",
    last_activity: "Oxirgi faollik",
    no_active_agents: "Bugun faol agentlar yo'q",
    products_catalog: "Mahsulotlar katalogi",
    search: "Qidiruv...",
    add_product: "Yangi Mahsulot Qo'shish",
    warehouse_stock: "Ombor qoldig'i",
    current_agent_stock: "Hozirda agentlarda mavjud",
    barcode: "Shtrix kod",
    product_name: "Mahsulot nomi",
    selling_price: "Sotish narxi",
    original_price: "Asl narxi",
    unit: "O'lchov birligi",
    stock_qty: "Ombor qoldig'i (Ombor)",
    assigned_agents: "Biriktirilgan Agentlar",
    total_assigned: "Jami berilgan (Olgan)",
    currently_remaining: "Hozirda qolgan",
    action: "Amal",
    not_entered: "Kiritilmagan",
    all_products_count: "Umumiy: {count} turdagi mahsulot",
    agents_count_title: "Umumiy: {count} turdagi mahsulot",
    stores_list: "Do'konlar ro'yxati",
    add_store: "Yangi Do'kon Qo'shish",
    import_excel: "Excel orqali yuklash",
    store_owner: "Rahbari",
    phone: "Telefon",
    address: "Manzil",
    geo: "Geologiya (Kenglik / Uzunlik)",
    route: "Yo'nalish",
    sde: "SDE",
    status: "Status",
    actions: "Harakatlar",
    edit: "Tahrirlash",
    delete: "O'chirish",
    normal: "Normal",
    error_ret: "Xato RET",
    agents_list: "Agentlar ro'yxati",
    add_agent: "Yangi Agent Qo'shish",
    agent_login: "Agent Logini",
    phone_number: "Telefon raqami",
    agent_route: "Mashrut (Yo'nalish)",
    assigned_products: "Biriktirilgan tovarlar",
    assigned_stores: "Biriktirilgan do'konlar",
    agent_status: "Holati",
    active: "Faol",
    inactive: "Nofaol",
    agent_assignments: "Agentlarga yuklamalar",
    assign_product: "Mahsulot biriktirish",
    assign_store: "Do'kon biriktirish",
    products_for_today: "Mahsulotlar (Bugun uchun)",
    stores_assigned: "Do'konlar (Biriktirilgan)",
    no_assignments: "Bugun uchun mahsulot biriktirilmagan.",
    no_stores_assigned: "Biriktirilgan do'konlar mavjud emas.",
    sales_history_title: "Sotuvlar tarixi",
    receipt: "Chek",
    date: "Sana",
    total_sum: "Jami summa",
    no_sales_found: "Ushbu muddat bo'yicha sotuvlar topilmadi.",
    cost: "Jami Tannarx (Xarajat)",
    top_agent: "Eng Yaxshi Agent",
    top_store: "Eng Faol Do'kon",
    top_product: "Ko'p Sotilgan Tovar",
    daily_sales_dynamics: "Kunlik savdo dinamikasi",
    daily_chart_desc: "Kunlik tushum va sof foyda diagrammasi",
    revenue_chart_label: "Tushum",
    profit_chart_label: "Sof foyda",
    bought_items_count: "Xarid qilgan tovarlar soni",
    best_selling_product: "Eng ko'p sotilgan tovar",
    serving_agents: "Xizmat ko'rsatgan agentlar",
    store_sales_modal_title: "{store} do'koniga sotilgan mahsulotlar",
    store_sales_modal_desc: "Sotuv hajmi eng yuqori bo'lgan mahsulotlar ro'yxati (eng ko'p tushum bo'yicha saralangan)",
    store_serving_agents_title: "Xizmat ko'rsatgan agentlar",
    select_agent_desc: "Iltimos, tahlil qilish uchun agentni tanlang",
    agent_sales: "Jami Savdo",
    agent_profit: "Sof Foyda",
    visited_stores: "Do'konlar",
    agent_sales_title: "{agent} Tahlili",
    agent_sales_desc: "Username: @{username} | Tel: {phone}",
    sold_products_details: "Sotilgan mahsulotlar tafsiloti",
    visited_stores_title: "Savdo qilingan do'konlar",
    no_sales_by_agent: "Bu agent hozircha savdo qilmagan.",
    no_stores_by_agent: "Ushbu agentga tegishli do'kon savdolari topilmadi.",
    profile_title: "Profil sozlamalari",
    profile_desc: "Tizim ma'muri (Admin) ma'lumotlarini o'zgartirish",
    admin_name_label: "Admin ismi",
    admin_photo_label: "Profil rasmi",
    upload_photo_btn: "Rasmni yuklash",
    save_btn: "Saqlash",
    save_success: "Profil ma'lumotlari muvaffaqiyatli saqlandi!",
    language_title: "Til sozlamalari",
    language_desc: "Dasturning interfeys tilini o'zgartirish",
    system_language: "Tizim tili",
    uzbek: "O'zbek tili",
    russian: "Rus tili (Русский)",
    cancel: "Bekor qilish",
    save: "Saqlash",
    close: "Yopish",
    product: "Mahsulot",
    quantity: "Miqdor",
    store: "Do'kon",
    dona: "dona",
    blok: "blok",
    quti: "quti",
    litr: "litr",
    kg: "kg",
    add_product_title: "Yangi mahsulot qo'shish",
    product_name_label: "Mahsulot nomi",
    price_label: "Sotish narxi",
    original_price_label: "Asl narxi (Tannarxi - Ixtiyoriy)",
    unit_label: "O'lchov birligi",
    initial_stock_label: "Boshlang'ich zaxira",
    barcode_label: "Shtrix-kod (Kiritish yoki skanerlash)",
    scanner_not_found: "Shtrix-kod topilmadi. Yangi mahsulot qo'shish oynasi ochildi.",
    found_scanner: "Topildi: {name} (Zaxira: {stock} {unit})",
    add_store_title: "Yangi do'kon qo'shish",
    owner_name_label: "Do'kon rahbari ismi",
    phone_label: "Telefon raqami",
    address_label: "Manzili",
    map_link_label: "Xarita havolasi (Google/Yandex Map link) - Kordinata olish uchun",
    lat_label: "Kenglik (Latitude)",
    lng_label: "Uzunlik (Longitude)",
    route_label: "Yo'nalish (Route)",
    status_label: "Status (Xato RET)",
    add_agent_title: "Yangi agent qo'shish",
    agent_name_label: "Agent ismi va familiyasi",
    agent_username_label: "Username (@ belgisisiz)",
    agent_phone_label: "Telefon raqami",
    agent_route_label: "Biriktiriladigan yo'nalishlar (Route)",
    agent_login_label: "Tizim uchun Login",
    agent_password_label: "Tizim uchun Parol",
    assign_product_title: "Mahsulot biriktirish",
    select_product_label: "Mahsulotni tanlang",
    qty_label: "Beriladigan miqdor",
    assign_store_title: "Do'kon biriktirish",
    select_store_label: "Do'konni tanlang",
    agent_stock_details_title: "Ushbu mahsulotning agentlardagi qoldiq tafsilotlari",
    agent_login_col: "Agent Logini",
    username_col: "Username",
    given_col: "Olgan",
    remaining_col: "Qolgan",
    no_username: "Username yo'q",
    store_sales_details_title: "{store} do'koniga sotilgan mahsulotlar",
    store_sales_details_desc: "Agent: {agent} (@{username}) tomonidan amalga oshirilgan savdolar tafsiloti",
    product_dist_title: "{product} tarqatilishi",
    product_dist_desc: "Tanlangan muddat bo'yicha mahsulotning do'konlarga tarqatilish va agentlar savdosi tafsilotlari",
    total_sold_badge: "JAMI SOTILGAN SONI",
    total_stores_badge: "TARQATILGAN DO'KONLAR SONI",
    tushum_badge: "UMUMIY TUSHUM",
    profit_badge: "UMUMIY SOF FOYDA",
    no_dist_found: "Ushbu mahsulot bo'yicha tarqatish ma'lumotlari topilmadi.",
    settings_discounts: "Chegirma sozlamalari",
    discounts_title: "Chegirma sozlamalari",
    discounts_desc: "Agentlar uchun sotuv vaqtida chegirma berish qoidalari va stavkalarini boshqarish",
    custom_discount_toggle_label: "Agentlarga erkin (ixtiyoriy) chegirma berishga ruxsat berish",
    custom_discount_toggle_desc: "Agar yoqilsa, agent to'lov qabul qilish vaqtida istalgan foizni yozishi mumkin. O'chirilsa, faqat quyida belgilangan tayyor chegirma foizlaridan birini tanlay oladi.",
    defined_discounts_title: "Belgilangan chegirma foizlari",
    add_discount_btn: "Chegirma foizini qo'shish",
    discount_percentage_label: "Chegirma foizi (%)",
    enter_discount_placeholder: "Masalan: 15",
    discount_exists_error: "Ushbu chegirma foizi allaqachon mavjud!",
    invalid_discount_error: "Foiz qiymati 1 va 99 oralig'ida bo'lishi kerak!",
    discount_added_success: "Chegirma foizi qo'shildi",
    discount_deleted_success: "Chegirma foizi o'chirildi"
  },
  ru: {
    login_title: "Вход в систему",
    login_subtitle: "Панель управления складом и агентами",
    username: "Имя пользователя",
    password: "Пароль",
    login_btn: "Войти",
    login_error: "Неверное имя пользователя или пароль!",
    demo_mode: "Демо: admin / admin (Полный админ) или ombor / ombor (Кладовщик)",
    dashboard: "Главная Панель",
    products: "Товары",
    stores: "Магазины",
    agents: "Управление агентами",
    analytics: "Аналитика",
    general_analytics: "Общий анализ",
    store_analytics: "Анализ магазинов",
    agent_analytics: "Анализ агентов",
    assignments: "Выдача агентам",
    my_tasks: "Мои задачи",
    sales_history: "История продаж",
    settings: "Настройки",
    profile_settings: "Настройки профиля",
    language_settings: "Настройки языка",
    logout: "Выход",
    active_in_system: "Активен в системе",
    date_filter: "Период анализа",
    bugun: "Сегодня",
    dan: "с",
    gacha: "по",
    custom_range: "Выбрать период",
    kechagi: "Вчера",
    "3-kunlik": "3 дня",
    "7-kunlik": "7 дней",
    "15-kunlik": "15 дней",
    "30-kunlik": "30 дней",
    hammasi: "Все время",
    revenue: "Выручка",
    active_agents: "Активные агенты",
    sold_items: "Продано товаров",
    net_profit: "Чистая прибыль",
    weekly_dynamics: "Динамика еженедельных продаж",
    active_agents_today: "Активные агенты на сегодня",
    agent_name: "Имя",
    stores_count: "Кол-во магазинов",
    total_sales: "Всего продаж",
    last_activity: "Последняя активность",
    no_active_agents: "Нет активных агентов на сегодня",
    products_catalog: "Каталог товаров",
    search: "Поиск...",
    add_product: "Добавить новый товар",
    warehouse_stock: "Остаток на складе",
    current_agent_stock: "Сейчас у агентов",
    barcode: "Штрих-код",
    product_name: "Название товара",
    selling_price: "Цена продажи",
    original_price: "Себестоимость",
    unit: "Ед. изм.",
    stock_qty: "Остаток (Склад)",
    assigned_agents: "Привязанные агенты",
    total_assigned: "Всего выдано",
    currently_remaining: "Осталось сейчас",
    action: "Действие",
    not_entered: "Не указано",
    all_products_count: "Всего: {count} видов товаров",
    agents_count_title: "Всего: {count} видов товаров",
    stores_list: "Список магазинов",
    add_store: "Добавить новый магазин",
    import_excel: "Импорт из Excel",
    store_owner: "Владелец",
    phone: "Телефон",
    address: "Адрес",
    geo: "Геолокация (Шир. / Долг.)",
    route: "Маршрут",
    sde: "SDE",
    status: "Статус",
    actions: "Действия",
    edit: "Редактировать",
    delete: "Удалить",
    normal: "Нормальный",
    error_ret: "Ошибка RET",
    agents_list: "Список агентов",
    add_agent: "Добавить нового агента",
    agent_login: "Логин агента",
    phone_number: "Номер телефона",
    agent_route: "Маршрут (Направление)",
    assigned_products: "Привязанные товары",
    assigned_stores: "Привязанные магазины",
    agent_status: "Статус",
    active: "Активный",
    inactive: "Неактивный",
    agent_assignments: "Выдача агентам",
    assign_product: "Привязать товар",
    assign_store: "Привязать магазин",
    products_for_today: "Товары (На сегодня)",
    stores_assigned: "Магазины (Привязанные)",
    no_assignments: "Товары на сегодня не привязаны.",
    no_stores_assigned: "Привязанные магазины отсутствуют.",
    sales_history_title: "История продаж",
    receipt: "Чек",
    date: "Дата",
    total_sum: "Общая сумма",
    no_sales_found: "Продажи за этот период не найдены.",
    cost: "Себестоимость",
    top_agent: "Лучший агент",
    top_store: "Активный магазин",
    top_product: "Популярный товар",
    daily_sales_dynamics: "Динамика ежедневных продаж",
    daily_chart_desc: "Диаграмма ежедневной выручки и чистой прибыли",
    revenue_chart_label: "Выручка",
    profit_chart_label: "Чистая прибыль",
    bought_items_count: "Куплено товаров",
    best_selling_product: "Самый продаваемый товар",
    serving_agents: "Обслуживающие агенты",
    store_sales_modal_title: "Проданные товары в магазин {store}",
    store_sales_modal_desc: "Список товаров с наибольшим объемом продаж (отсортировано по выручке)",
    store_serving_agents_title: "Обслуживающие агенты",
    select_agent_desc: "Пожалуйста, выберите агента для анализа",
    agent_sales: "Всего продаж",
    agent_profit: "Чистая прибыль",
    visited_stores: "Магазины",
    agent_sales_title: "Анализ {agent}",
    agent_sales_desc: "Имя пользователя: @{username} | Тел: {phone}",
    sold_products_details: "Детали проданных товаров",
    visited_stores_title: "Проданные магазины",
    no_sales_by_agent: "Этот агент еще не совершал продаж.",
    no_stores_by_agent: "Продажи этого агента по магазинам не найдены.",
    profile_title: "Настройки профиля",
    profile_desc: "Изменение данных администратора системы",
    admin_name_label: "Имя администратора",
    admin_photo_label: "Фото профиля",
    upload_photo_btn: "Загрузить фото",
    save_btn: "Сохранить",
    save_success: "Данные профиля успешно сохранены!",
    language_title: "Настройки языка",
    language_desc: "Изменение языка интерфейса приложения",
    system_language: "Язык системы",
    uzbek: "Узбекский язык (O'zbek)",
    russian: "Русский язык",
    cancel: "Отмена",
    save: "Сохранить",
    close: "Закрыть",
    product: "Товар",
    quantity: "Количество",
    store: "Магазин",
    dona: "шт",
    blok: "блок",
    quti: "кор",
    litr: "литр",
    kg: "кг",
    add_product_title: "Добавить новый товар",
    product_name_label: "Название товара",
    price_label: "Цена продажи",
    original_price_label: "Себестоимость (Необязательно)",
    unit_label: "Единица измерения",
    initial_stock_label: "Начальный запас",
    barcode_label: "Штрих-код (Ввод или сканирование)",
    scanner_not_found: "Штрих-код не найден. Открыто окно добавления нового товара.",
    found_scanner: "Найден: {name} (Запас: {stock} {unit})",
    add_store_title: "Добавить новый магазин",
    owner_name_label: "Имя владельца магазина",
    phone_label: "Номер телефона",
    address_label: "Адрес",
    map_link_label: "Ссылка на карту (Google/Yandex Map) - для получения координат",
    lat_label: "Широта (Latitude)",
    lng_label: "Долгота (Longitude)",
    route_label: "Маршрут (Route)",
    status_label: "Статус (Ошибка RET)",
    add_agent_title: "Добавить нового агента",
    agent_name_label: "Имя и фамилия агента",
    agent_username_label: "Имя пользователя (без @)",
    agent_phone_label: "Номер телефона",
    agent_route_label: "Назначаемые маршруты (Route)",
    agent_login_label: "Логин для входа",
    agent_password_label: "Пароль для входа",
    assign_product_title: "Привязать товар",
    select_product_label: "Выберите товар",
    qty_label: "Выдаваемое количество",
    assign_store_title: "Привязать магазин",
    select_store_label: "Выберите магазин",
    agent_stock_details_title: "Детали остатка этого товара у агентов",
    agent_login_col: "Логин агента",
    username_col: "Имя пользователя",
    given_col: "Выдано",
    remaining_col: "Осталось",
    no_username: "Нет юзернейма",
    store_sales_details_title: "Проданные товары в магазин {store}",
    store_sales_details_desc: "Агент: {agent} (@{username}) детали совершенных продаж",
    product_dist_title: "Распределение {product}",
    product_dist_desc: "Детали распределения товара по магазинам и продаж агентов за выбранный период",
    total_sold_badge: "ОБЩЕЕ КОЛ-ВО ПРОДАННОГО",
    total_stores_badge: "КОЛ-ВО РАСПРЕДЕЛЕННЫХ МАГАЗИНОВ",
    tushum_badge: "ОБЩАЯ ВЫРУЧКА",
    profit_badge: "ОБЩАЯ ЧИСТАЯ ПРИБЫЛЬ",
    no_dist_found: "Информации о распределении этого товара не найдено.",
    settings_discounts: "Настройки скидок",
    discounts_title: "Настройки скидок",
    discounts_desc: "Управление правилами и ставками предоставления скидок для агентов при продаже",
    custom_discount_toggle_label: "Разрешить агентам вводить произвольную скидку",
    custom_discount_toggle_desc: "Если включено, агент может ввести любой процент при приеме оплаты. Если выключено, он сможет выбрать только один из предопределенных процентов скидок.",
    defined_discounts_title: "Предопределенные проценты скидок",
    add_discount_btn: "Добавить процент скидки",
    discount_percentage_label: "Процент скидки (%)",
    enter_discount_placeholder: "Например: 15",
    discount_exists_error: "Этот процент скидки уже существует!",
    invalid_discount_error: "Значение процента должно быть от 1 до 99!",
    discount_added_success: "Процент скидки добавлен",
    discount_deleted_success: "Процент скидки удален"
  }
};

const initialProducts = [
  { id: 1, barcode: "48200001", name: "IQOS Iluma One (Pebble Grey)", price: 350000, originalPrice: 300000, unit: "dona", stock: 150, is_active: true },
  { id: 2, barcode: "48200002", name: "Heets Amber Selection", price: 18000, originalPrice: 15000, unit: "blok", stock: 1200, is_active: true },
  { id: 3, barcode: "48200003", name: "IQOS Terea Silver", price: 22000, originalPrice: 19000, unit: "blok", stock: 800, is_active: true },
  { id: 4, barcode: "48200004", name: "Fiit Regular", price: 17000, originalPrice: 14000, unit: "blok", stock: 650, is_active: true },
];

const initialStores = [
  { id: 14489, name: "G'ofur Ota Mini Market", owner_name: "G'ofurjon akam", phone: "+998 90 123 45 67", address: "Toshkent sh., Chilonzor 6-daha", map_link: "https://maps.google.com/?q=41.2842,69.1863", latitude: 41.2842, longitude: 69.1863, route: 'NM_01', sde: 'SDE_NM_01', census_type: 'Census    RET', status: 0 },
  { id: 57196, name: "Premium Smoke Shop", owner_name: "Davronbek", phone: "+998 93 543 21 09", address: "Toshkent sh., Amir Temur ko'chasi 12", map_link: "https://maps.google.com/?q=41.3113,69.2797", latitude: 41.3113, longitude: 69.2797, route: 'NM_01', sde: 'SDE_NM_01', census_type: 'Census    RET', status: 0 },
  { id: 53110, name: "24/7 Baza Do'kon", owner_name: "Azamat", phone: "+998 99 999 88 77", address: "Toshkent sh., Yunusobod 11-kvartal", map_link: "https://maps.google.com/?q=41.3654,69.2891", latitude: 41.3654, longitude: 69.2891, route: 'NM_02', sde: 'SDE_NM_02', census_type: 'Census    RET', status: 0 },
];

const initialAgents = [
  { id: 2, login: "AGENT-FG-R1", password: "123", username: "sherzod_agent", name: "Sherzod Alimov", role: "agent", phone: "+998 94 333 22 11", is_active: true },
  { id: 3, login: "AGENT-FG-R2", password: "123", username: "malika_agent", name: "Malika Qodirova", role: "agent", phone: "+998 97 777 55 44", is_active: true },
];

const getTodayDateString = (offsetDays = 0) => {
  const d = new Date();
  d.setDate(d.getDate() - offsetDays);
  return d.toISOString().split('T')[0];
};

const initialSales = [
  { 
    id: 101, 
    agentId: 2, 
    agent: "Sherzod Alimov", 
    storeId: 57196,
    store: "Premium Smoke Shop", 
    total: 1490000, 
    payment: "Payme", 
    status: "completed", 
    date: getTodayDateString(0), // Today
    items: [
      { productName: "IQOS Iluma One (Pebble Grey)", qty: 4, price: 350000, originalPrice: 300000 },
      { productName: "Heets Amber Selection", qty: 5, price: 18000, originalPrice: 15000 }
    ]
  },
  { 
    id: 102, 
    agentId: 3, 
    agent: "Malika Qodirova", 
    storeId: 14489,
    store: "G'ofur Ota Mini Market", 
    total: 610000, 
    payment: "Click", 
    status: "completed", 
    date: getTodayDateString(0), // Today
    items: [
      { productName: "IQOS Terea Silver", qty: 20, price: 22000, originalPrice: 19000 },
      { productName: "Fiit Regular", qty: 10, price: 17000, originalPrice: 14000 }
    ]
  },
  { 
    id: 103, 
    agentId: 2, 
    agent: "Sherzod Alimov", 
    storeId: 53110,
    store: "24/7 Baza Do'kon", 
    total: 2144000, 
    payment: "Payme", 
    status: "completed", 
    date: getTodayDateString(1), // Yesterday
    items: [
      { productName: "IQOS Iluma One (Pebble Grey)", qty: 5, price: 350000, originalPrice: 300000 },
      { productName: "Heets Amber Selection", qty: 20, price: 18000, originalPrice: 15000 },
      { productName: "Fiit Regular", qty: 2, price: 17000, originalPrice: 14000 }
    ]
  },
  { 
    id: 104, 
    agentId: 3, 
    agent: "Malika Qodirova", 
    storeId: 57196,
    store: "Premium Smoke Shop", 
    total: 850000, 
    payment: "Click", 
    status: "completed", 
    date: getTodayDateString(5), // 5 days ago
    items: [
      { productName: "Heets Amber Selection", qty: 30, price: 18000, originalPrice: 15000 },
      { productName: "Fiit Regular", qty: 10, price: 17000, originalPrice: 14000 }
    ]
  },
  { 
    id: 105, 
    agentId: 2, 
    agent: "Sherzod Alimov", 
    storeId: 14489,
    store: "G'ofur Ota Mini Market", 
    total: 1200000, 
    payment: "Payme", 
    status: "completed", 
    date: getTodayDateString(12), // 12 days ago
    items: [
      { productName: "IQOS Iluma One (Pebble Grey)", qty: 3, price: 350000, originalPrice: 300000 },
      { productName: "Fiit Regular", qty: 5, price: 17000, originalPrice: 14000 }
    ]
  },
  { 
    id: 106, 
    agentId: 3, 
    agent: "Malika Qodirova", 
    storeId: 53110,
    store: "24/7 Baza Do'kon", 
    total: 360000, 
    payment: "Click", 
    status: "completed", 
    date: getTodayDateString(25), // 25 days ago
    items: [
      { productName: "IQOS Terea Silver", qty: 10, price: 22000, originalPrice: 19000 },
      { productName: "Heets Amber Selection", qty: 5, price: 18000, originalPrice: 15000 }
    ]
  }
];

const chartData = [
  { name: 'Dush', savdo: 4200000 },
  { name: 'Sesh', savdo: 5100000 },
  { name: 'Chor', savdo: 4800000 },
  { name: 'Pay', savdo: 6200000 },
  { name: 'Jum', savdo: 7500000 },
  { name: 'Shan', savdo: 9100000 },
  { name: 'Yak', savdo: 8300000 },
];

function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [userRole, setUserRole] = useState(() => localStorage.getItem('userRole') || 'admin'); // admin, warehouse_manager, agent
  const [currentUserId, setCurrentUserId] = useState(() => localStorage.getItem('currentUserId') || '');
  const [activeTab, setActiveTab] = useState('dashboard');
  const [loginError, setLoginError] = useState('');

  // Interactive Lists
  const [products, setProducts] = useState(initialProducts);
  const [stores, setStores] = useState(initialStores);
  const [agents, setAgents] = useState(initialAgents);
  const [admins, setAdmins] = useState([]);
  const [newAdmin, setNewAdmin] = useState({ login: '', password: '', name: '', phone: '', role: 'admin' });
  const [editingAdmin, setEditingAdmin] = useState(null);
  const [showEditAdminModal, setShowEditAdminModal] = useState(false);
  const [sales, setSales] = useState(initialSales);
  const [assignments, setAssignments] = useState([
    { id: 1, agentId: 2, agentName: "Sherzod Alimov", productName: "IQOS Iluma One (Pebble Grey)", qty: 50, remainingQty: 50, date: new Date().toISOString().split('T')[0] },
    { id: 2, agentId: 2, agentName: "Sherzod Alimov", productName: "Heets Amber Selection", qty: 200, remainingQty: 200, date: new Date().toISOString().split('T')[0] },
    { id: 3, agentId: 3, agentName: "Malika Qodirova", productName: "IQOS Terea Silver", qty: 150, remainingQty: 150, date: new Date().toISOString().split('T')[0] },
  ]);

  const [storeAssignments, setStoreAssignments] = useState([
    { id: 1, agentId: 2, agentName: "Sherzod Alimov", storeName: "Premium Smoke Shop", ownerName: "Davronbek", phone: "+998 93 543 21 09", address: "Toshkent sh., Amir Temur ko'chasi 12", date: new Date().toISOString().split('T')[0], order: 1 },
    { id: 2, agentId: 2, agentName: "Sherzod Alimov", storeName: "24/7 Baza Do'kon", ownerName: "Azamat", phone: "+998 99 999 88 77", address: "Toshkent sh., Yunusobod 11-kvartal", date: new Date().toISOString().split('T')[0], order: 2 },
    { id: 3, agentId: 3, agentName: "Malika Qodirova", storeName: "G'ofur Ota Mini Market", ownerName: "G'ofurjon akam", phone: "+998 90 123 45 67", address: "Toshkent sh., Chilonzor 6-daha", date: new Date().toISOString().split('T')[0], order: 1 },
  ]);

  const [selectedAgentId, setSelectedAgentId] = useState(2); // Defaults to Sherzod Alimov (id: 2)
  const [productsTabMode, setProductsTabMode] = useState('warehouse'); // 'warehouse' or 'agents_stock'
  const [selectedProductStockDetails, setSelectedProductStockDetails] = useState(null); // stores product name for detailed stock modal
  const [tahlillarOpen, setTahlillarOpen] = useState(false);
  const [selectedTahlilAgentId, setSelectedTahlilAgentId] = useState(2);
  const [selectedTahlilStoreDetails, setSelectedTahlilStoreDetails] = useState(null);
  const [selectedTahlilAgentStore, setSelectedTahlilAgentStore] = useState(null);
  const [selectedTahlilProductDetails, setSelectedTahlilProductDetails] = useState(null);
  const [dateFilter, setDateFilter] = useState('bugun');
  const [startDate, setStartDate] = useState(() => getTodayDateString(0));
  const [endDate, setEndDate] = useState(() => getTodayDateString(0));

  const [settingsOpen, setSettingsOpen] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const API_URL = 'https://posagent.onrender.com/api';
  const [isLoading, setIsLoading] = useState(false);
  const [token, setToken] = useState(() => localStorage.getItem('token') || '');

  useEffect(() => {
    setIsMobileMenuOpen(false);
  }, [activeTab]);

  const loadCloudData = async (authToken) => {
    if (!authToken) return;
    setIsLoading(true);
    try {
      // 1. Fetch Products
      const productsRes = await fetch(`${API_URL}/products`, {
        headers: { 'Authorization': `Bearer ${authToken}` }
      });
      if (!productsRes.ok) throw new Error('Mahsulotlarni yuklashda xatolik');
      const productsData = await productsRes.json();

      // 2. Fetch Stores
      const storesRes = await fetch(`${API_URL}/stores`, {
        headers: { 'Authorization': `Bearer ${authToken}` }
      });
      if (!storesRes.ok) throw new Error('Do\'konlarni yuklashda xatolik');
      const storesData = await storesRes.json();

      // 3. Fetch Agents (Users) - Only for admin/warehouse_manager
      let agentsData = [];
      const storedRole = localStorage.getItem('userRole') || userRole;
      if (storedRole !== 'agent') {
        const agentsRes = await fetch(`${API_URL}/auth/users`, {
          headers: { 'Authorization': `Bearer ${authToken}` }
        });
        if (!agentsRes.ok) throw new Error('Agentlarni yuklashda xatolik');
        agentsData = await agentsRes.json();
      }

      // 4. Fetch Sales
      const isAgent = (storedRole === 'agent');
      const salesUrl = isAgent 
        ? `${API_URL}/sales/agent/${localStorage.getItem('currentUserId') || currentUserId}`
        : `${API_URL}/sales`;
        
      const salesRes = await fetch(salesUrl, {
        headers: { 'Authorization': `Bearer ${authToken}` }
      });
      if (!salesRes.ok) throw new Error('Savdolarni yuklashda xatolik');
      const salesData = await salesRes.json();

      // Map products
      const mappedProducts = productsData.map(p => ({
        id: p.id,
        barcode: p.barcode,
        name: p.name,
        price: parseFloat(p.price),
        originalPrice: parseFloat(p.original_price || 0),
        unit: p.unit,
        stock: p.stock,
        is_active: p.is_active
      }));

      // Map stores
      const mappedStores = storesData.map(s => ({
        id: s.id,
        name: s.name,
        owner_name: s.owner_name,
        phone: s.phone,
        address: s.address,
        map_link: s.map_link,
        latitude: parseFloat(s.location_lat || 0),
        longitude: parseFloat(s.location_lng || 0),
        route: s.route || 'NM_01',
        sde: s.sde || 'SDE_NM_01',
        census_type: s.census_type || 'Census RET',
        status: s.status || 0,
        agentId: s.agent_id,
        assigned_date: s.assigned_date,
        duration_days: s.duration_days,
        order: s.order
      }));

      // Map agents
      const mappedAgents = agentsData.filter(u => u.role === 'agent').map(u => ({
        id: u.id,
        login: u.username,
        username: u.username,
        name: u.name || u.username,
        role: u.role,
        phone: u.phone || '',
        is_active: u.is_active
      }));

      // Map sales
      const mappedSales = salesData.map(sale => ({
        id: sale.id,
        agentId: sale.agent_id,
        agent: sale.agent ? sale.agent.name : 'Noma\'lum',
        agentPhone: sale.agent ? sale.agent.phone : '',
        agentUsername: sale.agent ? sale.agent.username : '',
        storeId: sale.store_id,
        store: sale.store ? sale.store.name : 'Noma\'lum',
        total: parseFloat(sale.total_amount),
        payment: sale.transaction ? sale.transaction.payment_gateway : 'Naqd',
        status: sale.status,
        date: sale.createdAt.split('T')[0],
        items: (sale.items || []).map(item => ({
          productName: item.product ? item.product.name : 'Noma\'lum',
          qty: item.quantity,
          price: parseFloat(item.unit_price),
          originalPrice: parseFloat(item.original_price || 0)
        }))
      }));

      // Map admins
      const mappedAdmins = agentsData.filter(u => u.role === 'admin' || u.role === 'warehouse_manager').map(u => ({
        id: u.id,
        login: u.username,
        username: u.username,
        name: u.name || u.username,
        role: u.role,
        phone: u.phone || '',
        is_active: u.is_active
      }));

      setProducts(mappedProducts);
      setStores(mappedStores);
      setAgents(mappedAgents);
      setAdmins(mappedAdmins);
      setSales(mappedSales);

      // 5. Dynamic Store Assignments loading
      const loadedStoreAssignments = mappedStores
        .filter(s => s.agentId !== null && s.agentId !== undefined)
        .map(s => {
          const agent = mappedAgents.find(a => a.id === s.agentId) || { name: 'Agent' };
          return {
            id: s.id,
            agentId: s.agentId,
            agentName: agent.name || agent.login,
            storeName: s.name,
            ownerName: s.owner_name,
            phone: s.phone,
            address: s.address,
            map_link: s.map_link,
            latitude: s.latitude,
            longitude: s.longitude,
            date: s.assigned_date || new Date().toISOString().split('T')[0],
            durationDays: s.duration_days || 1,
            order: s.order || 1
          };
        })
        .sort((a, b) => (a.order || 0) - (b.order || 0));
      setStoreAssignments(loadedStoreAssignments);

      // 6. Fetch Product Assignments (Agent Inventory)
      const activeAgentId = storedRole === 'agent' 
        ? parseInt(localStorage.getItem('currentUserId') || currentUserId || '0')
        : selectedAgentId;

      if (activeAgentId) {
        try {
          const inventoryRes = await fetch(`${API_URL}/inventory/agent/${activeAgentId}`, {
            headers: { 'Authorization': `Bearer ${authToken}` }
          });
          if (inventoryRes.ok) {
            const inventoryData = await inventoryRes.json();
            const mappedAssignments = inventoryData.map(item => ({
              id: item.id,
              agentId: item.agent_id,
              agentName: (mappedAgents.find(a => a.id === item.agent_id) || { name: 'Agent' }).name,
              productId: item.product_id,
              productName: item.product ? item.product.name : 'Noma\'lum',
              price: item.product ? item.product.price : 0,
              barcode: item.product ? item.product.barcode : '',
              unit: item.product ? item.product.unit : 'dona',
              qty: item.qty_given,
              remainingQty: item.qty_given - item.qty_sold - item.qty_returned,
              date: item.date,
              durationDays: item.duration_days || 1
            }));
            setAssignments(mappedAssignments);
          }
        } catch (err) {
          console.error("Failed to fetch product assignments:", err);
        }
      }

    } catch (error) {
      console.error(error);
      showAlert(error.message, 'error');
      if (error.message.includes('401') || error.message.includes('403') || error.message.includes('token') || error.message.includes('sessiya')) {
        handleLogout();
      }
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = () => {
    setToken('');
    setCurrentUserId('');
    localStorage.removeItem('token');
    localStorage.removeItem('userRole');
    localStorage.removeItem('currentUserId');
    localStorage.removeItem('adminName');
    localStorage.removeItem('adminPhoto');
    setIsLoggedIn(false);
    showAlert('Tizimdan chiqildi', 'info');
  };

  useEffect(() => {
    if (token) {
      setIsLoggedIn(true);
      loadCloudData(token);
    }
  }, [token]);

  // Global responsive tables handler for mobile WebViews without :has() support
  useEffect(() => {
    const tables = document.querySelectorAll('table');
    tables.forEach(table => {
      const parent = table.parentElement;
      if (parent && parent.tagName === 'DIV') {
        if (!parent.classList.contains('table-responsive-container')) {
          parent.classList.add('table-responsive-container');
        }
      }
    });
  });

  // Fetch product assignments when selectedAgentId changes (for admin view)
  useEffect(() => {
    if (token && selectedAgentId && userRole !== 'agent') {
      const fetchInventory = async () => {
        try {
          const res = await fetch(`${API_URL}/inventory/agent/${selectedAgentId}`, {
            headers: { 'Authorization': `Bearer ${token}` }
          });
          if (res.ok) {
            const data = await res.json();
            const mapped = data.map(item => ({
              id: item.id,
              agentId: item.agent_id,
              agentName: (agents.find(a => a.id === item.agent_id) || { name: 'Agent' }).name,
              productId: item.product_id,
              productName: item.product ? item.product.name : 'Noma\'lum',
              price: item.product ? item.product.price : 0,
              barcode: item.product ? item.product.barcode : '',
              unit: item.product ? item.product.unit : 'dona',
              qty: item.qty_given,
              remainingQty: item.qty_given - item.qty_sold - item.qty_returned,
              date: item.date
            }));
            setAssignments(mapped);
          }
        } catch (err) {
          console.error("Error fetching selected agent inventory:", err);
        }
      };
      fetchInventory();
    }
  }, [selectedAgentId, token, userRole, agents]);

  const [language, setLanguage] = useState(() => localStorage.getItem('lang') || 'uz');
  const [adminName, setAdminName] = useState(() => localStorage.getItem('adminName') || 'Bosh Admin');
  const [adminPhoto, setAdminPhoto] = useState(() => localStorage.getItem('adminPhoto') || '');
  const [theme, setTheme] = useState(() => localStorage.getItem('theme') || 'dark');
  const [customDiscountEnabled, setCustomDiscountEnabled] = useState(() => {
    const stored = localStorage.getItem('customDiscountEnabled');
    return stored ? stored === 'true' : true;
  });
  const [discountsList, setDiscountsList] = useState(() => {
    const stored = localStorage.getItem('discountsList');
    return stored ? JSON.parse(stored) : [10, 20, 50];
  });

  // Agent Cashier (POS) and Payment Integration States
  const [activeCashierStore, setActiveCashierStore] = useState(null);
  const [cashierCart, setCashierCart] = useState([]);
  const [barcodeInput, setBarcodeInput] = useState('');
  const [searchProductQuery, setSearchProductQuery] = useState('');
  const [selectedPaymentMethod, setSelectedPaymentMethod] = useState('Naqd');
  const [cashierDiscount, setCashierDiscount] = useState(0);
  const [customDiscountInput, setCustomDiscountInput] = useState('');
  const [showPaymentSection, setShowPaymentSection] = useState(false);
  const [paymentIntegrations, setPaymentIntegrations] = useState(() => {
    const stored = localStorage.getItem('payment_integrations');
    return stored ? JSON.parse(stored) : ['Naqd', 'Click', 'Payme'];
  });
  const [showProductModal, setShowProductModal] = useState(false);
  const [selectedModalProducts, setSelectedModalProducts] = useState([]);
  const [showCameraScanner, setShowCameraScanner] = useState(false);
  const [modalSearchQuery, setModalSearchQuery] = useState('');

  // Tinda Terminal Settings States
  const [tindaTerminalIp, setTindaTerminalIp] = useState(() => localStorage.getItem('tinda_terminal_ip') || '192.168.1.100:8080');
  const [tindaTerminalLogin, setTindaTerminalLogin] = useState(() => localStorage.getItem('tinda_terminal_login') || '');
  const [tindaTerminalPin, setTindaTerminalPin] = useState(() => localStorage.getItem('tinda_terminal_pin') || '');
  const [tindaDefaultMxik, setTindaDefaultMxik] = useState(() => localStorage.getItem('tinda_default_mxik') || '09901001001000000');
  const [tindaDefaultPackage, setTindaDefaultPackage] = useState(() => localStorage.getItem('tinda_default_package') || '242030');

  // Tinda Transaction Live States
  const [tindaPaymentStatus, setTindaPaymentStatus] = useState(null); // 'connecting', 'logging_in', 'waiting_card', 'success', 'error'
  const [tindaErrorMessage, setTindaErrorMessage] = useState('');
  const [tindaSocket, setTindaSocket] = useState(null);
  const [showCashierTerminalConfig, setShowCashierTerminalConfig] = useState(false);
  const [visitedStores, setVisitedStores] = useState(() => {
    const stored = localStorage.getItem('visited_stores');
    return stored ? JSON.parse(stored) : [];
  });
  const [showExitQuestionnaire, setShowExitQuestionnaire] = useState(false);
  const [exitReason, setExitReason] = useState('');

  // Company Branding States (Persisted in localStorage)
  const [companyLogo, setCompanyLogo] = useState(() => localStorage.getItem('companyLogo') || null);
  const [companyName, setCompanyName] = useState(() => localStorage.getItem('companyName') || 'AGENT POS');
  const [companyNameColor, setCompanyNameColor] = useState(() => localStorage.getItem('companyNameColor') || '#f8fafc');
  const [companyBio, setCompanyBio] = useState(() => localStorage.getItem('companyBio') || 'MANAGEMENT');
  const [companyBioColor, setCompanyBioColor] = useState(() => localStorage.getItem('companyBioColor') || '#0d9488');
  const isAssignmentActive = (assignDateStr, durationDays = 1) => {
    if (!assignDateStr) return false;
    if (durationDays === 9999 || durationDays === 0) return true;
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const assignDate = new Date(assignDateStr);
    assignDate.setHours(0, 0, 0, 0);

    const diffTime = today.getTime() - assignDate.getTime();
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays >= 0 && diffDays < durationDays;
  };

  // Selected Agent and their assigned Products/Stores for everywhere in the App (reactive)
  const activeAgent = userRole === 'agent' 
    ? { id: parseInt(currentUserId || localStorage.getItem('currentUserId') || '0'), login: adminName, username: adminName } 
    : agents.find(a => a.id === selectedAgentId);

  const agentProducts = activeAgent 
    ? assignments.filter(ass => ass.agentId === activeAgent.id && isAssignmentActive(ass.date, ass.durationDays || 1)) 
    : [];

  const agentStores = activeAgent 
    ? storeAssignments.filter(ass => ass.agentId === activeAgent.id && isAssignmentActive(ass.date, ass.durationDays || 1)).sort((a, b) => (a.order || 0) - (b.order || 0)) 
    : [];

  const activeAgentStores = agentStores.filter(store => {
    const todayStr = new Date().toISOString().split('T')[0];
    return !visitedStores.some(v => v.storeId === store.id && v.date === todayStr);
  });


  // Custom alert and confirm states
  const [toast, setToast] = useState(null); // { message: '', type: 'success' | 'error' | 'info' }
  const [confirmDialog, setConfirmDialog] = useState(null); // { message: '', onConfirm: () => void, onCancel: () => void }

  const showAlert = (message, type = 'info') => {
    setToast({ message, type });
    setTimeout(() => {
      setToast(prev => prev && prev.message === message ? null : prev);
    }, 4000);
  };

  const showConfirm = (message, onConfirm) => {
    setConfirmDialog({
      message,
      onConfirm: () => {
        onConfirm();
        setConfirmDialog(null);
      },
      onCancel: () => {
        setConfirmDialog(null);
      }
    });
  };

  const recordSoldVisit = (store, cart) => {
    if (!store) return;
    const newVisit = {
      storeId: store.id,
      storeName: store.storeName,
      status: 'sold',
      reason: '',
      items: cart.map(item => ({
        productName: item.productName || "Mahsulot",
        qty: item.quantity,
        price: item.price
      })),
      date: new Date().toISOString().split('T')[0],
      time: new Date().toLocaleTimeString('uz-UZ', { hour: '2-digit', minute: '2-digit' })
    };
    setVisitedStores(prev => {
      const updated = [newVisit, ...prev];
      localStorage.setItem('visited_stores', JSON.stringify(updated));
      return updated;
    });
  };

  const handleExitCashier = () => {
    if (userRole === 'agent') {
      setShowExitQuestionnaire(true);
    } else {
      showConfirm(
        language === 'uz' ? "Kassadan chiqishni xohlaysizmi? Savatcha tozalanadi!" : "Вы действительно хотите выйти из кассы? Корзина будет очищена!",
        () => {
          setActiveCashierStore(null);
          setCashierCart([]);
        }
      );
    }
  };

  const getAgentLoginByName = (agentName) => {
    if (!agentName) return '-';
    const found = agents.find(a => a.name === agentName || a.login === agentName);
    return found ? found.login : agentName;
  };

  const getAgentLoginById = (agentId) => {
    const found = agents.find(a => a.id === agentId);
    return found ? found.login : '-';
  };

  const t = (key, params = {}) => {
    const lang = language || 'uz';
    let text = translations[lang]?.[key] || translations['uz']?.[key] || key;
    Object.keys(params).forEach(p => {
      text = text.replace(`{${p}}`, params[p]);
    });
    return text;
  };

  // Form States & Modals
  const [showAddProductModal, setShowAddProductModal] = useState(false);
  const [showAddStoreModal, setShowAddStoreModal] = useState(false);
  const [showAssignProductModal, setShowAssignProductModal] = useState(false);
  const [showAssignStoreModal, setShowAssignStoreModal] = useState(false);
  const [showEditAssignmentModal, setShowEditAssignmentModal] = useState(false);
  const [selectedAssignmentForEdit, setSelectedAssignmentForEdit] = useState(null);
  const [editAssignmentQty, setEditAssignmentQty] = useState('');
  const [editAssignmentRemainingQty, setEditAssignmentRemainingQty] = useState('');

  const [newProduct, setNewProduct] = useState({ barcode: '', name: '', price: '', originalPrice: '', unit: 'dona', stock: '' });
  const [productSearchQuery, setProductSearchQuery] = useState('');
  const [scannerValue, setScannerValue] = useState('');
  const [scannerNotification, setScannerNotification] = useState('');

  const [newStore, setNewStore] = useState({
    id: '',
    name: '',
    owner_name: '',
    phone: '',
    address: '',
    map_link: '',
    latitude: '',
    longitude: '',
    route: 'NM_01',
    sde: 'SDE_NM_01',
    census_type: 'Census    RET',
    status: '0'
  });
  const [editingStore, setEditingStore] = useState(null);
  const [showEditStoreModal, setShowEditStoreModal] = useState(false);
  const [selectedStoreIds, setSelectedStoreIds] = useState([]);
  const [newAssignment, setNewAssignment] = useState({ agentId: '', productId: '', qty: '', durationDays: '1', isPermanent: false });
  const [newStoreAssignment, setNewStoreAssignment] = useState({ agentId: '', storeId: '', durationDays: '1', isPermanent: false });

  const [newAgent, setNewAgent] = useState({
    login: 'AGENT-FG-R3',
    password: '',
    username: '',
    name: '',
    phone: ''
  });
  const [editingAgent, setEditingAgent] = useState(null);
  const [showEditAgentModal, setShowEditAgentModal] = useState(false);

  useEffect(() => {
    resetNewAgentForm();
  }, []);

  useEffect(() => {
    if (theme === 'light') {
      document.documentElement.classList.add('light');
    } else {
      document.documentElement.classList.remove('light');
    }
    localStorage.setItem('theme', theme);
  }, [theme]);

  const handleLogin = async (e) => {
    e.preventDefault();
    if (!username || !password) {
      setLoginError('Login va parolni kiriting');
      return;
    }
    
    setIsLoading(true);
    setLoginError('');
    try {
      const res = await fetch(`${API_URL}/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Kirishda xatolik yuz berdi');
      }

      if (data.user.role !== 'admin' && data.user.role !== 'warehouse_manager' && data.user.role !== 'agent') {
        throw new Error('Ushbu panelga kirishga ruxsatingiz yo\'q!');
      }

      setToken(data.token);
      localStorage.setItem('token', data.token);
      setUserRole(data.user.role);
      localStorage.setItem('userRole', data.user.role);
      setCurrentUserId(data.user.id);
      localStorage.setItem('currentUserId', data.user.id);
      setAdminName(data.user.name || data.user.username);
      localStorage.setItem('adminName', data.user.name || data.user.username);
      setIsLoggedIn(true);
      
      let defaultTab = 'products';
      if (data.user.role === 'admin') {
        defaultTab = 'dashboard';
      } else if (data.user.role === 'agent') {
        defaultTab = 'assignments';
      }
      setActiveTab(defaultTab);
      showAlert('Tizimga muvaffaqiyatli kirildi!', 'success');
    } catch (error) {
      setLoginError(error.message);
    } finally {
      setIsLoading(false);
    }
  };

  const getFilteredSales = (salesList) => {
    const todayStr = getTodayDateString(0);
    const yesterdayStr = getTodayDateString(1);

    return salesList.filter(sale => {
      const saleDate = sale.date; 
      let normalizedDate = saleDate;
      if (saleDate === 'Bugun') {
        normalizedDate = todayStr;
      } else if (saleDate === 'Kecha') {
        normalizedDate = yesterdayStr;
      }

      if (dateFilter === 'bugun') {
        return normalizedDate === todayStr;
      }

      if (dateFilter === 'custom') {
        return (!startDate || normalizedDate >= startDate) && (!endDate || normalizedDate <= endDate);
      }

      return true;
    });
  };

  const handleAddProduct = async (e) => {
    e.preventDefault();
    if (!newProduct.barcode || !newProduct.name || !newProduct.price) return;
    
    const sellPrice = parseFloat(newProduct.price);
    const costPrice = newProduct.originalPrice ? parseFloat(newProduct.originalPrice) : 0;

    if (costPrice > 0 && sellPrice <= costPrice) {
      showAlert(
        language === 'uz' 
          ? "Sotish narxi asl narxidan (tannarxidan) katta bo'lishi shart! Zarariga sotish mumkin emas." 
          : "Цена продажи должна быть больше себестоимости! Нельзя продавать в убыток.", 
        'error'
      );
      return;
    }
    
    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/products`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          barcode: newProduct.barcode,
          name: newProduct.name,
          price: parseFloat(newProduct.price),
          original_price: newProduct.originalPrice ? parseFloat(newProduct.originalPrice) : 0.00,
          unit: newProduct.unit,
          stock: parseInt(newProduct.stock) || 0
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Mahsulot qo\'shishda xatolik yuz berdi');
      }

      showAlert('Mahsulot muvaffaqiyatli qo\'shildi!', 'success');
      setNewProduct({ barcode: '', name: '', price: '', originalPrice: '', unit: 'dona', stock: '' });
      setShowAddProductModal(false);
      await loadCloudData(token);
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleScannerSubmit = (e) => {
    e.preventDefault();
    if (!scannerValue) return;

    const foundProduct = products.find(p => p.barcode === scannerValue);
    if (foundProduct) {
      setScannerNotification(`Topildi: ${foundProduct.name} (Zaxira: ${foundProduct.stock} ${foundProduct.unit})`);
      setProductSearchQuery(scannerValue);
      setTimeout(() => setScannerNotification(''), 5000);
    } else {
      setNewProduct(prev => ({ ...prev, barcode: scannerValue }));
      setShowAddProductModal(true);
      setScannerNotification(`Shtrix-kod topilmadi. Yangi mahsulot qo'shish oynasi ochildi.`);
      setTimeout(() => setScannerNotification(''), 5000);
    }
    setScannerValue('');
  };

  const handleUpdateProductStock = (productId, val) => {
    setProducts(prevProducts => prevProducts.map(p => {
      if (p.id === productId) {
        return {
          ...p,
          stock: val === "" ? "" : (isNaN(parseInt(val)) ? 0 : parseInt(val))
        };
      }
      return p;
    }));
  };

  const resolveShortUrl = (url) => {
    return new Promise((resolve) => {
      try {
        const https = window.require('https');
        https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, (res) => {
          if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
            resolve(res.headers.location);
          } else {
            resolve(url);
          }
        }).on('error', () => {
          resolve(url);
        });
      } catch (e) {
        resolve(url);
      }
    });
  };

  const parseCoordinatesFromUrl = (url) => {
    if (!url) return null;

    // Google Maps marker coordinates in data parameter (most precise)
    const googleMarkerRegex = /!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/;
    const matchGoogleMarker = url.match(googleMarkerRegex);
    if (matchGoogleMarker) {
      return {
        lat: parseFloat(matchGoogleMarker[1]),
        lng: parseFloat(matchGoogleMarker[2])
      };
    }
    
    // Google Maps @lat,lng
    const googleAtRegex = /@(-?\d+\.\d+),(-?\d+\.\d+)/;
    const matchGoogleAt = url.match(googleAtRegex);
    if (matchGoogleAt) {
      return {
        lat: parseFloat(matchGoogleAt[1]),
        lng: parseFloat(matchGoogleAt[2])
      };
    }

    try {
      // If full URL with searchParams
      const urlObj = new URL(url);
      const searchParams = urlObj.searchParams;

      // Google maps q=lat,lng
      if (searchParams.has('q')) {
        const q = searchParams.get('q');
        const parts = q.split(',');
        if (parts.length === 2) {
          const lat = parseFloat(parts[0]);
          const lng = parseFloat(parts[1]);
          if (!isNaN(lat) && !isNaN(lng)) {
            return { lat, lng };
          }
        }
      }

      // Yandex maps ll=lng,lat
      if (searchParams.has('ll')) {
        const ll = searchParams.get('ll');
        const parts = ll.split(',');
        if (parts.length === 2) {
          const lng = parseFloat(parts[0]);
          const lat = parseFloat(parts[1]);
          if (!isNaN(lat) && !isNaN(lng)) {
            return { lat, lng };
          }
        }
      }

      // Yandex maps whatshere[point]=lng,lat
      if (searchParams.has('whatshere[point]')) {
        const pt = searchParams.get('whatshere[point]');
        const parts = pt.split(',');
        if (parts.length === 2) {
          const lng = parseFloat(parts[0]);
          const lat = parseFloat(parts[1]);
          if (!isNaN(lat) && !isNaN(lng)) {
            return { lat, lng };
          }
        }
      }
    } catch (e) {
      // Parse error
    }

    // Fallback: Generic regex for coordinates "lat,lng" or "lng,lat"
    const coordsRegex = /(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/;
    const matchCoords = url.match(coordsRegex);
    if (matchCoords) {
      const val1 = parseFloat(matchCoords[1]);
      const val2 = parseFloat(matchCoords[2]);
      if (val1 > 50 && val2 < 50) {
        return { lat: val2, lng: val1 };
      }
      return { lat: val1, lng: val2 };
    }

    return null;
  };

  const handleMapLinkChange = async (url, isEditing = false) => {
    // 1. Update the link input immediately so the user sees what they pasted
    if (isEditing) {
      setEditingStore(prev => ({ ...prev, map_link: url }));
    } else {
      setNewStore(prev => ({ ...prev, map_link: url }));
    }

    if (!url) return;

    // 2. Try parsing coordinates directly from the pasted URL
    let coords = parseCoordinatesFromUrl(url);
    if (coords) {
      if (isEditing) {
        setEditingStore(prev => ({
          ...prev,
          map_link: url,
          latitude: coords.lat,
          longitude: coords.lng
        }));
      } else {
        setNewStore(prev => ({
          ...prev,
          map_link: url,
          latitude: coords.lat,
          longitude: coords.lng
        }));
      }
      return;
    }

    // 3. If no coordinates in the URL, check if it's a short URL we can resolve
    if (url.includes('maps.app.goo.gl') || url.includes('/maps/-')) {
      try {
        const resolvedUrl = await resolveShortUrl(url);
        if (resolvedUrl) {
          const resolvedCoords = parseCoordinatesFromUrl(resolvedUrl);
          if (resolvedCoords) {
            if (isEditing) {
              setEditingStore(prev => ({
                ...prev,
                latitude: resolvedCoords.lat,
                longitude: resolvedCoords.lng
              }));
            } else {
              setNewStore(prev => ({
                ...prev,
                latitude: resolvedCoords.lat,
                longitude: resolvedCoords.lng
              }));
            }
          }
        }
      } catch (err) {
        console.error("Short URL resolution failed:", err);
      }
    }
  };

  const handleAddStore = async (e) => {
    e.preventDefault();
    if (!newStore.name || !newStore.address) return;
    
    const latVal = parseFloat(newStore.latitude) || '';
    const lngVal = parseFloat(newStore.longitude) || '';

    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/stores`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          name: newStore.name,
          owner_name: newStore.owner_name || 'Tadbirkor',
          phone: newStore.phone || '+998 90 000 00 00',
          address: newStore.address,
          map_link: latVal && lngVal ? `https://maps.google.com/?q=${latVal},${lngVal}` : (newStore.map_link || `https://maps.google.com/?q=${newStore.name}`),
          location_lat: latVal.toString(),
          location_lng: lngVal.toString()
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Do\'kon qo\'shishda xatolik yuz berdi');
      }

      showAlert('Do\'kon muvaffaqiyatli qo\'shildi!', 'success');
      setNewStore({
        id: '',
        name: '',
        owner_name: '',
        phone: '',
        address: '',
        map_link: '',
        latitude: '',
        longitude: '',
        route: 'NM_01',
        sde: 'SDE_NM_01',
        census_type: 'Census    RET',
        status: '0'
      });
      setShowAddStoreModal(false);
      await loadCloudData(token);
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleUpdateStore = async (e) => {
    e.preventDefault();
    if (!editingStore.name || !editingStore.address) return;

    const latVal = parseFloat(editingStore.latitude) || '';
    const lngVal = parseFloat(editingStore.longitude) || '';

    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/stores/${editingStore.id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          name: editingStore.name,
          owner_name: editingStore.owner_name,
          phone: editingStore.phone,
          address: editingStore.address,
          map_link: latVal && lngVal ? `https://maps.google.com/?q=${latVal},${lngVal}` : (editingStore.map_link || `https://maps.google.com/?q=${editingStore.name}`),
          location_lat: latVal.toString(),
          location_lng: lngVal.toString()
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Do\'konni yangilashda xatolik yuz berdi');
      }

      showAlert('Do\'kon muvaffaqiyatli yangilandi!', 'success');
      setEditingStore(null);
      setShowEditStoreModal(false);
      await loadCloudData(token);
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteStore = (id) => {
    showConfirm(
      language === 'uz' ? "Ushbu do'konni o'chirishni tasdiqlaysizmi?" : "Вы действительно хотите удалить этот магазин?",
      async () => {
        setIsLoading(true);
        try {
          const res = await fetch(`${API_URL}/stores/${id}`, {
            method: 'DELETE',
            headers: {
              'Authorization': `Bearer ${token}`
            }
          });

          const data = await res.json();
          if (!res.ok) {
            throw new Error(data.error || 'Do\'konni o\'chirishda xatolik yuz berdi');
          }

          showAlert(
            language === 'uz' ? "Do'kon muvaffaqiyatli o'chirildi" : "Магазин успешно удален",
            'success'
          );
          setSelectedStoreIds(selectedStoreIds.filter(sid => sid !== id));
          await loadCloudData(token);
        } catch (error) {
          showAlert(error.message, 'error');
        } finally {
          setIsLoading(false);
        }
      }
    );
  };

  const handleDeleteSelectedStores = () => {
    if (selectedStoreIds.length === 0) return;
    showConfirm(
      language === 'uz' 
        ? `Tanlangan ${selectedStoreIds.length} ta do'konni o'chirishni tasdiqlaysizmi?` 
        : `Вы действительно хотите удалить выбранные магазины (${selectedStoreIds.length} шт)?`,
      async () => {
        setIsLoading(true);
        try {
          for (const id of selectedStoreIds) {
            await fetch(`${API_URL}/stores/${id}`, {
              method: 'DELETE',
              headers: {
                'Authorization': `Bearer ${token}`
              }
            });
          }
          showAlert(
            language === 'uz' ? "Tanlangan do'konlar o'chirildi" : "Выбранные магазины удалены",
            'success'
          );
          setSelectedStoreIds([]);
          await loadCloudData(token);
        } catch (error) {
          showAlert(error.message, 'error');
        } finally {
          setIsLoading(false);
        }
      }
    );
  };

  const resetNewAgentForm = (currentAgents = agents) => {
    let maxNum = 0;
    currentAgents.forEach(agent => {
      if (agent.login && agent.login.startsWith('AGENT-FG-R')) {
        const numPart = agent.login.replace('AGENT-FG-R', '');
        const num = parseInt(numPart);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    });
    setNewAgent({
      login: `AGENT-FG-R${maxNum + 1}`,
      password: '',
      username: '',
      name: '',
      phone: ''
    });
  };

  const handleAddAgent = async (e) => {
    e.preventDefault();
    if (!newAgent.login || !newAgent.password) {
      showAlert(
        language === 'uz' ? "Login va parolni kiritish majburiy!" : "Логин и пароль обязательны!", 
        'error'
      );
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/auth/register`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          username: newAgent.login,
          password: newAgent.password,
          name: newAgent.name || '',
          phone: newAgent.phone || '',
          role: 'agent'
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Agent qo\'shishda xatolik yuz berdi');
      }

      showAlert('Agent muvaffaqiyatli qo\'shildi!', 'success');
      await loadCloudData(token);
      resetNewAgentForm();
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleUpdateAgent = async (e) => {
    e.preventDefault();
    if (!editingAgent.login || !editingAgent.password) {
      showAlert(
        language === 'uz' ? "Login va parolni kiritish majburiy!" : "Логин и пароль обязательны!", 
        'error'
      );
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/auth/users/${editingAgent.id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          username: editingAgent.login,
          password: editingAgent.password,
          name: editingAgent.name || '',
          phone: editingAgent.phone || '',
          is_active: editingAgent.is_active
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Agentni yangilashda xatolik yuz berdi');
      }

      showAlert('Agent muvaffaqiyatli yangilandi!', 'success');
      setShowEditAgentModal(false);
      setEditingAgent(null);
      await loadCloudData(token);
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteAgent = (id) => {
    showConfirm(
      language === 'uz' ? "Ushbu agentni tizimdan o'chirishni tasdiqlaysizmi?" : "Вы действительно хотите удалить этого агента?",
      async () => {
        setIsLoading(true);
        try {
          const res = await fetch(`${API_URL}/auth/users/${id}`, {
            method: 'DELETE',
            headers: {
              'Authorization': `Bearer ${token}`
            }
          });

          const data = await res.json();
          if (!res.ok) {
            throw new Error(data.error || 'Agentni o\'chirishda xatolik yuz berdi');
          }

          showAlert('Agent muvaffaqiyatli o\'chirildi!', 'success');
          await loadCloudData(token);
        } catch (error) {
          showAlert(error.message, 'error');
        } finally {
          setIsLoading(false);
        }
      }
    );
  };

  const handleAddAdmin = async (e) => {
    e.preventDefault();
    if (!newAdmin.login || !newAdmin.password) {
      showAlert(
        language === 'uz' ? "Login va parolni kiritish majburiy!" : "Логин и пароль обязательны!", 
        'error'
      );
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/auth/register`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          username: newAdmin.login,
          password: newAdmin.password,
          name: newAdmin.name || '',
          phone: newAdmin.phone || '',
          role: newAdmin.role || 'admin'
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Admin qo\'shishda xatolik yuz berdi');
      }

      showAlert(
        language === 'uz' ? 'Yangi administrator muvaffaqiyatli qo\'shildi!' : 'Новый администратор успешно добавлен!',
        'success'
      );
      setNewAdmin({ login: '', password: '', name: '', phone: '', role: 'admin' });
      await loadCloudData(token);
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleUpdateAdmin = async (e) => {
    e.preventDefault();
    if (!editingAdmin.login) {
      showAlert(
        language === 'uz' ? "Login kiritish majburiy!" : "Логин обязателен!", 
        'error'
      );
      return;
    }

    setIsLoading(true);
    try {
      const res = await fetch(`${API_URL}/auth/users/${editingAdmin.id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          username: editingAdmin.login,
          password: editingAdmin.password || undefined,
          name: editingAdmin.name || '',
          phone: editingAdmin.phone || '',
          role: editingAdmin.role || 'admin',
          is_active: editingAdmin.is_active
        })
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Adminni yangilashda xatolik yuz berdi');
      }

      showAlert(
        language === 'uz' ? 'Administrator ma\'lumotlari yangilandi!' : 'Данные администратора обновлены!',
        'success'
      );
      setShowEditAdminModal(false);
      setEditingAdmin(null);
      await loadCloudData(token);
    } catch (error) {
      showAlert(error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDeleteAdmin = (id) => {
    showConfirm(
      language === 'uz' ? "Ushbu administratorni tizimdan o'chirishni tasdiqlaysizmi?" : "Вы действительно хотите удалить этого администратора?",
      async () => {
        setIsLoading(true);
        try {
          const res = await fetch(`${API_URL}/auth/users/${id}`, {
            method: 'DELETE',
            headers: {
              'Authorization': `Bearer ${token}`
            }
          });

          const data = await res.json();
          if (!res.ok) {
            throw new Error(data.error || 'Adminni o\'chirishda xatolik yuz berdi');
          }

          showAlert(
            language === 'uz' ? 'Administrator o\'chirildi!' : 'Администратор удален!',
            'success'
          );
          await loadCloudData(token);
        } catch (error) {
          showAlert(error.message, 'error');
        } finally {
          setIsLoading(false);
        }
      }
    );
  };

  const handleExportExcel = () => {
    try {
      const headers = ["Идентификатор", "Торговая точка", "Широта", "Долгота", "ret", "Сотрудник", "Census", "Некорректный RET"];
      const rows = [headers];

      stores.forEach(store => {
        rows.push([
          store.id,
          store.name,
          store.latitude || '',
          store.longitude || '',
          store.route || 'NM_01',
          store.sde || `SDE_${store.route || 'NM_01'}`,
          store.census_type || 'Census    RET',
          store.status || 0
        ]);
      });

      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Do'konlar");
      
      XLSX.writeFile(wb, "NM_1_Exported.xlsx");
      showAlert(
        language === 'uz' ? "Do'konlar muvaffaqiyatli Excel fayliga yuklab olindi (NM_1_Exported.xlsx)!" : "Магазины успешно экспортированы в Excel (NM_1_Exported.xlsx)!", 
        'success'
      );
    } catch (err) {
      console.error("Excel export error:", err);
      showAlert(
        language === 'uz' ? "Xatolik: Excel fayliga yozishda xatolik yuz berdi." : "Ошибка: Не удалось экспортировать в Excel.", 
        'error'
      );
    }
  };

  const handleImportExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        
        const importedStores = [];
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (row && row[1]) {
            const id = row[0] || (Date.now() + i);
            const rawName = String(row[1]);
            const cleanName = rawName.replace(/["']/g, '').split('(')[0].trim();
            const lat = row[2] || '';
            const lng = row[3] || '';
            const route = row[4] || 'NM_01';
            
            importedStores.push({
              id: parseInt(id) || (Date.now() + i),
              name: cleanName,
              owner_name: "Tadbirkor",
              phone: "+998 90 000 00 00",
              address: `Namangan, yo'nalish: ${route}`,
              latitude: parseFloat(lat) || '',
              longitude: parseFloat(lng) || '',
              route: route,
              sde: row[5] || `SDE_${route}`,
              census_type: row[6] || 'Census    RET',
              status: parseInt(row[7]) || 0,
              map_link: lat && lng ? `https://maps.google.com/?q=${lat},${lng}` : `https://maps.google.com/?q=${cleanName}`
            });
          }
        }

        if (importedStores.length > 0) {
          (async () => {
            setIsLoading(true);
            try {
              const res = await fetch(`${API_URL}/stores/bulk`, {
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json',
                  'Authorization': `Bearer ${token}`
                },
                body: JSON.stringify(importedStores)
              });

              const data = await res.json();
              if (!res.ok) {
                throw new Error(data.error || 'Do\'konlarni bulutga yuklashda xatolik yuz berdi');
              }

              showAlert(
                language === 'uz' 
                  ? `Muvaffaqiyatli yuklandi! Jami ${importedStores.length} ta do'kon bulutli ma'lumotlar bazasiga saqlandi.` 
                  : `Успешно загружено! Всего ${importedStores.length} магазинов сохранено в облачную базу данных.`, 
                'success'
              );
              await loadCloudData(token);
            } catch (error) {
              console.error("Bulk upload error:", error);
              showAlert(error.message, 'error');
            } finally {
              setIsLoading(false);
            }
          })();
        }
      } catch (err) {
        console.error("Excel import error:", err);
        showAlert(
          language === 'uz' ? "Xatolik: Excel faylini o'qib bo'lmadi. Formatini tekshiring." : "Ошибка: Не удалось прочитать файл Excel. Проверьте формат.", 
          'error'
        );
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null;
  };

  const handleAddAssignment = (e) => {
    e.preventDefault();
    if (!newAssignment.productId || !newAssignment.qty) return;
    
    const agent = agents.find(a => a.id === selectedAgentId);
    const product = products.find(p => p.id === parseInt(newAssignment.productId));

    if (!agent || !product) return;

    const requestedQty = parseInt(newAssignment.qty);
    if (product.stock < requestedQty) {
      showAlert(
        language === 'uz' 
          ? `Omborda yetarli miqdor yo'q! Hozirgi qoldiq: ${product.stock} ${product.unit}` 
          : `Недостаточно товара на складе! Текущий остаток: ${product.stock} ${product.unit}`, 
        'error'
      );
      return;
    }

    // Call backend API to save the assignment
    fetch(`${API_URL}/inventory/assign`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        agent_id: agent.id,
        date: new Date().toISOString().split('T')[0],
        duration_days: newAssignment.isPermanent ? 9999 : parseInt(newAssignment.durationDays || '1'),
        products: [
          {
            product_id: product.id,
            qty_given: requestedQty
          }
        ]
      })
    })
    .then(res => {
      if (!res.ok) throw new Error('Serverga saqlashda xatolik');
      return res.json();
    })
    .then(() => {
      showAlert(
        language === 'uz' ? 'Mahsulot muvaffaqiyatli biriktirildi' : 'Товар успешно закреплен',
        'success'
      );
      // Deduct stock from warehouse products list locally
      setProducts(products.map(p => {
        if (p.id === product.id) {
          return { ...p, stock: p.stock - requestedQty };
        }
        return p;
      }));

      // Reload assignments
      loadCloudData(token);
    })
    .catch(err => {
      console.error(err);
      showAlert(language === 'uz' ? 'Xatolik yuz berdi' : 'Произошла ошибка', 'error');
    });

    setNewAssignment({ agentId: '', productId: '', qty: '', durationDays: '1', isPermanent: false });
    setShowAssignProductModal(false);
  };

  const handleUpdateAssignment = (id, newQty, newRemainingQty) => {
    const ass = assignments.find(a => a.id === id);
    if (!ass) return false;

    const qtyDiff = newQty - ass.qty;
    const p = products.find(prod => prod.name === ass.productName);

    if (qtyDiff > 0 && p && p.stock < qtyDiff) {
      showAlert(
        language === 'uz'
          ? `Omborda yetarli mahsulot yo'q! Qo'shimcha so'ralgan: ${qtyDiff}, omborda: ${p.stock}`
          : `Недостаточно товара на складе! Дополнительно запрошено: ${qtyDiff}, на складе: ${p.stock}`,
        'error'
      );
      return false;
    }

    // Call backend API to update assignment
    fetch(`${API_URL}/inventory/${id}`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        qty_given: newQty
      })
    })
    .then(res => {
      if (!res.ok) {
        return res.json().then(err => { throw new Error(err.error || 'Yangilashda xatolik'); });
      }
      return res.json();
    })
    .then(() => {
      showAlert(
        language === 'uz'
          ? "Biriktirilgan mahsulot muvaffaqiyatli yangilandi"
          : "Закрепленный товар успешно обновлен",
        'success'
      );
      loadCloudData(token);
    })
    .catch(err => {
      console.error(err);
      showAlert(err.message, 'error');
    });

    return true;
  };

  const handleUpdateRemainingQty = (id, newQty) => {
    setAssignments(prev => prev.map(ass => {
      if (ass.id === id) {
        const oldRemainingVal = ass.remainingQty !== undefined ? ass.remainingQty : ass.qty;
        const oldRemaining = oldRemainingVal === "" ? 0 : oldRemainingVal;

        let capped;
        if (newQty === "") {
          capped = "";
        } else {
          capped = Math.max(0, Math.min(ass.qty, newQty));
        }

        const numericNewQty = capped === "" ? 0 : capped;
        const diff = oldRemaining - numericNewQty;

        if (diff !== 0) {
          setProducts(prevProducts => prevProducts.map(p => {
            if (p.name === ass.productName) {
              return { ...p, stock: p.stock + diff };
            }
            return p;
          }));
        }

        return { ...ass, remainingQty: capped };
      }
      return ass;
    }));
  };

  const handleAddStoreAssignment = (e) => {
    e.preventDefault();
    if (!newStoreAssignment.storeId) return;

    const agent = agents.find(a => a.id === selectedAgentId);
    const store = stores.find(s => s.id === parseInt(newStoreAssignment.storeId));

    if (!agent || !store) return;

    const nextOrder = storeAssignments.filter(ass => ass.agentId === agent.id).length + 1;

    // Send PUT request to save the assignment in the database
    fetch(`${API_URL}/stores/${store.id}`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify({
        name: store.name,
        owner_name: store.ownerName || store.owner_name,
        phone: store.phone,
        address: store.address,
        map_link: store.map_link,
        location_lat: String(store.latitude || store.location_lat || ''),
        location_lng: String(store.longitude || store.location_lng || ''),
        agent_id: agent.id,
        assigned_date: new Date().toISOString().split('T')[0],
        duration_days: newStoreAssignment.isPermanent ? 9999 : parseInt(newStoreAssignment.durationDays || '1'),
        order: nextOrder
      })
    })
    .then(res => {
      if (!res.ok) throw new Error('Do\'konni biriktirishda xatolik');
      return res.json();
    })
    .then(() => {
      showAlert(
        language === 'uz' ? 'Do\'kon muvaffaqiyatli biriktirildi' : 'Магазин успешно закреплен',
        'success'
      );
      // Reload all data to refresh state
      loadCloudData(token);
    })
    .catch(err => {
      console.error(err);
      showAlert(err.message, 'error');
    });

    setNewStoreAssignment({ agentId: '', storeId: '', durationDays: '1', isPermanent: false });
    setShowAssignStoreModal(false);
  };

  const [draggedStoreIndex, setDraggedStoreIndex] = useState(null);

  const handleStoreDragStart = (e, index) => {
    setDraggedStoreIndex(index);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleStoreDragOver = (e) => {
    e.preventDefault();
  };

  const handleStoreDrop = (e, targetIndex) => {
    e.preventDefault();
    if (draggedStoreIndex === null || draggedStoreIndex === targetIndex) return;

    const agent = agents.find(a => a.id === selectedAgentId);
    if (!agent) return;

    // Filter and sort by current order
    const agentStores = storeAssignments
      .filter(ass => ass.agentId === agent.id)
      .sort((a, b) => (a.order || 0) - (b.order || 0));

    // Reorder
    const updated = [...agentStores];
    const [moved] = updated.splice(draggedStoreIndex, 1);
    updated.splice(targetIndex, 0, moved);

    // Reassign order
    const updatedWithOrder = updated.map((item, idx) => ({
      ...item,
      order: idx + 1
    }));

    // Update main state locally
    setStoreAssignments(prev => {
      const others = prev.filter(ass => ass.agentId !== agent.id);
      return [...others, ...updatedWithOrder];
    });

    setDraggedStoreIndex(null);

    // Persist new orders to backend
    Promise.all(updatedWithOrder.map(item => {
      const fullStore = stores.find(s => s.id === item.id);
      if (!fullStore) return Promise.resolve();
      return fetch(`${API_URL}/stores/${item.id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          name: fullStore.name,
          owner_name: fullStore.ownerName || fullStore.owner_name,
          phone: fullStore.phone,
          address: fullStore.address,
          map_link: fullStore.map_link,
          location_lat: String(fullStore.latitude || fullStore.location_lat || ''),
          location_lng: String(fullStore.longitude || fullStore.location_lng || ''),
          agent_id: agent.id,
          order: item.order
        })
      });
    }))
    .then(() => {
      loadCloudData(token);
    })
    .catch(err => console.error("Drop reorder save error:", err));
  };

  const handleMoveStore = (index, direction) => {
    const targetIndex = index + direction;
    const agent = agents.find(a => a.id === selectedAgentId);
    if (!agent) return;

    // Filter and sort by current order
    const agentStores = storeAssignments
      .filter(ass => ass.agentId === agent.id)
      .sort((a, b) => (a.order || 0) - (b.order || 0));

    if (targetIndex < 0 || targetIndex >= agentStores.length) return;

    // Reorder
    const updated = [...agentStores];
    const [moved] = updated.splice(index, 1);
    updated.splice(targetIndex, 0, moved);

    // Reassign order
    const updatedWithOrder = updated.map((item, idx) => ({
      ...item,
      order: idx + 1
    }));

    // Update main state locally
    setStoreAssignments(prev => {
      const others = prev.filter(ass => ass.agentId !== agent.id);
      return [...others, ...updatedWithOrder];
    });

    // Persist new orders to backend
    Promise.all(updatedWithOrder.map(item => {
      const fullStore = stores.find(s => s.id === item.id);
      if (!fullStore) return Promise.resolve();
      return fetch(`${API_URL}/stores/${item.id}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          name: fullStore.name,
          owner_name: fullStore.ownerName || fullStore.owner_name,
          phone: fullStore.phone,
          address: fullStore.address,
          map_link: fullStore.map_link,
          location_lat: String(fullStore.latitude || fullStore.location_lat || ''),
          location_lng: String(fullStore.longitude || fullStore.location_lng || ''),
          agent_id: agent.id,
          order: item.order
        })
      });
    }))
    .then(() => {
      loadCloudData(token);
    })
    .catch(err => console.error("Move reorder save error:", err));
  };

  const handleDeleteStoreAssignment = (id) => {
    showConfirm(
      language === 'uz'
        ? "Ushbu do'konni agent biriktiruvidan o'chirishni xohlaysizmi?"
        : "Вы действительно хотите удалить привязку этого магазина?",
      () => {
        const store = stores.find(s => s.id === id);
        if (!store) return;

        fetch(`${API_URL}/stores/${id}`, {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
          },
          body: JSON.stringify({
            name: store.name,
            owner_name: store.ownerName || store.owner_name,
            phone: store.phone,
            address: store.address,
            map_link: store.map_link,
            location_lat: String(store.latitude || store.location_lat || ''),
            location_lng: String(store.longitude || store.location_lng || ''),
            agent_id: null,
            order: null
          })
        })
        .then(res => {
          if (!res.ok) throw new Error('Biriktiruvni o\'chirishda xatolik');
          return res.json();
        })
        .then(() => {
          showAlert(
            language === 'uz' ? "Do'kon biriktiruvi o'chirildi" : "Привязка магазина удалена",
            'success'
          );
          loadCloudData(token);
        })
        .catch(err => {
          console.error(err);
          showAlert(err.message, 'error');
        });
      }
    );
  };

  const handleDeleteProductAssignment = (id) => {
    const ass = assignments.find(a => a.id === id);
    if (!ass) return;

    showConfirm(
      language === 'uz'
        ? "Ushbu mahsulotni agent biriktiruvidan o'chirishni va qoldiqni omborga qaytarishni xohlaysizmi?"
        : "Вы действительно хотите удалить привязку этого товара и вернуть остаток на склад?",
      () => {
        // Call backend API to delete the assignment
        fetch(`${API_URL}/inventory/${id}`, {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${token}`
          }
        })
        .then(res => {
          if (!res.ok) throw new Error('Biriktiruvni o\'chirishda xatolik');
          return res.json();
        })
        .then(() => {
          const remainingVal = ass.remainingQty !== undefined ? ass.remainingQty : ass.qty;
          const remainingQtyToReturn = remainingVal === "" ? 0 : remainingVal;

          if (remainingQtyToReturn > 0) {
            setProducts(prevProducts => prevProducts.map(p => {
              if (p.name === ass.productName) {
                return { ...p, stock: p.stock + remainingQtyToReturn };
              }
              return p;
            }));
          }

          setAssignments(prev => prev.filter(a => a.id !== id));
          showAlert(
            language === 'uz' ? "Mahsulot biriktiruvi o'chirildi" : "Привязка товара удалена",
            'success'
          );
        })
        .catch(err => {
          console.error(err);
          showAlert(err.message, 'error');
        });
      }
    );
  };

  const handleDeleteProduct = (productId) => {
    const product = products.find(p => p.id === productId);
    if (!product) return;

    showConfirm(
      language === 'uz'
        ? `Haqiqatan ham "${product.name}" mahsulotini ombordan o'chirmoqchimisiz?`
        : `Вы действительно хотите удалить товар "${product.name}" со склада?`,
      async () => {
        setIsLoading(true);
        try {
          const res = await fetch(`${API_URL}/products/${productId}`, {
            method: 'DELETE',
            headers: {
              'Authorization': `Bearer ${token}`
            }
          });
          
          const data = await res.json();
          if (!res.ok) {
            throw new Error(data.error || 'Mahsulotni o\'chirishda xatolik yuz berdi');
          }

          showAlert(
            language === 'uz' ? "Mahsulot muvaffaqiyatli o'chirildi" : "Товар успешно удален",
            'success'
          );
          await loadCloudData(token);
        } catch (error) {
          showAlert(error.message, 'error');
        } finally {
          setIsLoading(false);
        }
      }
    );
  };

  const handleAddDiscount = (val) => {
    if (isNaN(val) || val < 1 || val > 99) {
      showAlert(t('invalid_discount_error'), 'error');
      return;
    }
    if (discountsList.includes(val)) {
      showAlert(t('discount_exists_error'), 'error');
      return;
    }
    const newList = [...discountsList, val].sort((a, b) => a - b);
    setDiscountsList(newList);
    localStorage.setItem('discountsList', JSON.stringify(newList));
    showAlert(t('discount_added_success'), 'success');
    
    const input = document.getElementById('new-discount-input');
    if (input) input.value = '';
  };

  const handleDeleteDiscount = (val) => {
    const newList = discountsList.filter(d => d !== val);
    setDiscountsList(newList);
    localStorage.setItem('discountsList', JSON.stringify(newList));
    showAlert(t('discount_deleted_success'), 'success');
  };

  // --- Agent Cashier Helper Functions ---

  const handleOpenCashier = (store) => {
    setActiveCashierStore(store);
    setCashierCart([]);
    setBarcodeInput('');
    setSearchProductQuery('');
    setCashierDiscount(0);
    setCustomDiscountInput('');
    setShowPaymentSection(false);
    // Load integrated payments
    const stored = localStorage.getItem('payment_integrations');
    const integrations = stored ? JSON.parse(stored) : ['Naqd', 'Click', 'Payme'];
    if (integrations.length > 0) {
      setSelectedPaymentMethod(integrations[0]);
    } else {
      setSelectedPaymentMethod('');
    }
  };

  const handleAddProductToCart = (product) => {
    // Determine available quantity
    const maxQty = product.qty - (product.qty_sold || 0);
    if (maxQty <= 0) {
      showAlert(language === 'uz' ? "Ushbu mahsulot qoldig'i tugagan!" : "Этот товар закончился!", 'error');
      return;
    }

    setCashierCart(prev => {
      const existing = prev.find(item => item.productId === product.productId || item.productId === product.id);
      if (existing) {
        if (existing.quantity >= maxQty) {
          showAlert(language === 'uz' ? `Omborda yetarli qoldiq yo'q! Maksimal: ${maxQty} dona` : `Недостаточно остатка на складе! Максимум: ${maxQty} шт`, 'error');
          return prev;
        }
        return prev.map(item => 
          (item.productId === product.productId || item.productId === product.id)
            ? { ...item, quantity: item.quantity + 1 }
            : item
        );
      } else {
        const price = product.price || 0;
        return [...prev, {
          productId: product.productId || product.id,
          productName: product.productName || product.name,
          barcode: product.barcode || '',
          price: price,
          quantity: 1,
          unit: product.unit || 'dona',
          maxQty: maxQty
        }];
      }
    });
  };

  const handleUpdateCartItemQty = (productId, newQty) => {
    const qty = parseInt(newQty);
    if (isNaN(qty) || qty <= 0) return;

    setCashierCart(prev => {
      const existing = prev.find(item => item.productId === productId);
      if (!existing) return prev;

      if (qty > existing.maxQty) {
        showAlert(language === 'uz' ? `Omborda yetarli qoldiq yo'q! Maksimal: ${existing.maxQty} dona` : `Недостаточно остатка на складе! Максимум: ${existing.maxQty} шт`, 'error');
        return prev.map(item => item.productId === productId ? { ...item, quantity: existing.maxQty } : item);
      }
      return prev.map(item => item.productId === productId ? { ...item, quantity: qty } : item);
    });
  };

  const handleRemoveCartItem = (productId) => {
    setCashierCart(prev => prev.filter(item => item.productId !== productId));
  };

  const handleBarcodeSubmit = (e) => {
    if (e) e.preventDefault();
    if (!barcodeInput.trim()) return;

    // Find product in agentProducts by barcode
    const found = agentProducts.find(p => p.barcode === barcodeInput.trim() || p.productBarcode === barcodeInput.trim());
    if (found) {
      handleAddProductToCart(found);
      setBarcodeInput('');
    } else {
      showAlert(language === 'uz' ? "Ushbu shtrix-kodga ega mahsulot topilmadi!" : "Товар с таким штрих-кодом не найден!", 'error');
    }
  };

  const handleCreateCashierSale = async () => {
    if (cashierCart.length === 0) {
      showAlert(language === 'uz' ? "Savatcha bo'sh!" : "Корзина пуста!", 'error');
      return;
    }

    if (!selectedPaymentMethod) {
      showAlert(language === 'uz' ? "Iltimos, to'lov turini tanlang!" : "Пожалуйста, выберите способ оплаты!", 'error');
      return;
    }

    try {
      let pct = 0;
      if (customDiscountEnabled && customDiscountInput) {
        pct = parseFloat(customDiscountInput) || 0;
      } else {
        pct = cashierDiscount;
      }

      const items = cashierCart.map(item => {
        const discountedPrice = Math.round(item.price * (1 - pct / 100));
        return {
          product_id: item.productId,
          quantity: item.quantity,
          unit_price: discountedPrice
        };
      });

      const payload = {
        store_id: activeCashierStore.id,
        payment_gateway: selectedPaymentMethod.toLowerCase(),
        items: items
      };

      const res = await fetch(`${API_URL}/sales`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + token
        },
        body: JSON.stringify(payload)
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Sale failed');
      }

      showAlert(language === 'uz' ? "Sotuv muvaffaqiyatli yakunlandi!" : "Продажа успешно завершена!", 'success');
      
      await loadCloudData(token);

      recordSoldVisit(activeCashierStore, cashierCart);
      setActiveCashierStore(null);
      setCashierCart([]);
      setShowPaymentSection(false);

    } catch (err) {
      showAlert(err.message, 'error');
    }
  };

  const handleTindaPayment = (subtotal, discountAmount, finalTotal) => {
    if (!tindaTerminalIp) {
      showAlert(language === 'uz' ? "Terminal IP manzili sozlanmagan!" : "IP-адрес терминала не настроен!", 'error');
      return;
    }

    setTindaPaymentStatus('connecting');
    setTindaErrorMessage('');

    const wsUrl = `ws://${tindaTerminalIp}`;
    const socket = new WebSocket(wsUrl);
    setTindaSocket(socket);

    // Timeout if cannot connect in 7 seconds
    const connectionTimeout = setTimeout(() => {
      if (socket.readyState !== WebSocket.OPEN) {
        socket.close();
        setTindaPaymentStatus('error');
        setTindaErrorMessage(language === 'uz' ? "Terminal bilan bog'lanib bo'lmadi! IP manzil va Wi-Fi ulanishini tekshiring." : "Не удалось подключиться к терминалу! Проверьте IP-адрес и подключение Wi-Fi.");
      }
    }, 7000);

    socket.onopen = () => {
      clearTimeout(connectionTimeout);
      
      // Step 2: Send Login Command
      setTindaPaymentStatus('logging_in');
      const loginRequest = {
        id: "login_" + Date.now(),
        command: "Login",
        data: {
          login: tindaTerminalLogin || "Admin",
          password: tindaTerminalPin || "1111"
        }
      };
      socket.send(JSON.stringify(loginRequest));
    };

    socket.onmessage = (event) => {
      try {
        const response = JSON.parse(event.data);
        console.log("Tinda Terminal Response:", response);

        if (response.command === "Login") {
          if (response.status === "Success" || response.error === "UserAlreadyLogined") {
            // Step 3: Send Sale Command
            setTindaPaymentStatus('waiting_card');
            
            const productsListForTinda = cashierCart.map(item => {
              return {
                id: String(item.productId),
                name: item.productName || "Mahsulot",
                barcode: item.barcode || null,
                units: item.unit || "dona",
                price: Math.round(item.price * 100),
                vat: 0.0,
                amount: parseFloat(item.quantity),
                isDecimalUnits: item.unit !== 'dona',
                psid: tindaDefaultMxik,
                packageCode: parseInt(tindaDefaultPackage),
                sellerType: 0
              };
            });

            const saleRequest = {
              id: "sale_" + Date.now(),
              command: "Sale",
              data: {
                payments: [
                  {
                    amount: Math.round(finalTotal * 100),
                    paymentType: "Card"
                  }
                ],
                amount: Math.round(subtotal * 100),
                discountAmount: Math.round(discountAmount * 100),
                totalAmount: Math.round(finalTotal * 100),
                saleType: "Sale",
                products: productsListForTinda,
                shouldPrintReceipt: true
              }
            };
            socket.send(JSON.stringify(saleRequest));
          } else {
            socket.close();
            setTindaPaymentStatus('error');
            setTindaErrorMessage(`${language === 'uz' ? 'Terminalga kirish xatosi' : 'Ошибка авторизации на терминале'}: ${response.error || 'Noma\'lum xato'}`);
          }
        } else if (response.command === "Sale") {
          socket.close();
          if (response.status === "Success") {
            setTindaPaymentStatus('success');
            setTimeout(() => {
              handleCompleteLocalSaleAfterTindaPayment(finalTotal, response.payload?.saleId);
            }, 1500);
          } else {
            setTindaPaymentStatus('error');
            setTindaErrorMessage(`${language === 'uz' ? 'To\'lov xatosi' : 'Ошибка оплаты'}: ${response.error || 'Noma\'lum xato'}`);
          }
        }
      } catch (err) {
        console.error("Error parsing socket message:", err);
      }
    };

    socket.onerror = (err) => {
      clearTimeout(connectionTimeout);
      socket.close();
      setTindaPaymentStatus('error');
      setTindaErrorMessage(language === 'uz' ? "Tarmoq xatosi yuz berdi!" : "Произошла ошибка сети!");
    };

    socket.onclose = () => {
      console.log("Tinda Terminal Socket Closed.");
    };
  };

  const handleCompleteLocalSaleAfterTindaPayment = async (finalTotal, tindaSaleId) => {
    try {
      let pct = 0;
      if (customDiscountEnabled && customDiscountInput) {
        pct = parseFloat(customDiscountInput) || 0;
      } else {
        pct = cashierDiscount;
      }

      const items = cashierCart.map(item => {
        const discountedPrice = Math.round(item.price * (1 - pct / 100));
        return {
          product_id: item.productId,
          quantity: item.quantity,
          unit_price: discountedPrice
        };
      });

      const payload = {
        store_id: activeCashierStore.id,
        payment_gateway: 'tinda',
        items: items
      };

      const res = await fetch(`${API_URL}/sales`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + token
        },
        body: JSON.stringify(payload)
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Sale failed');
      }

      showAlert(language === 'uz' ? "Sotuv muvaffaqiyatli yakunlandi!" : "Продажа успешно завершена!", 'success');
      
      await loadCloudData(token);

      recordSoldVisit(activeCashierStore, cashierCart);
      setActiveCashierStore(null);
      setCashierCart([]);
      setShowPaymentSection(false);
      setTindaPaymentStatus(null);

    } catch (err) {
      showAlert(err.message, 'error');
      setTindaPaymentStatus('error');
      setTindaErrorMessage(language === 'uz' ? "To'lov o'tdi, lekin bazaga saqlashda xatolik." : "Оплата прошла, но ошибка сохранения в БД.");
    }
  };

  // Camera scanner instance ref
  const scannerRef = React.useRef(null);

  const playBeep = () => {
    try {
      const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
      const oscillator = audioCtx.createOscillator();
      const gainNode = audioCtx.createGain();
      oscillator.connect(gainNode);
      gainNode.connect(audioCtx.destination);
      oscillator.type = 'sine';
      oscillator.frequency.setValueAtTime(800, audioCtx.currentTime);
      gainNode.gain.setValueAtTime(0.15, audioCtx.currentTime);
      oscillator.start();
      oscillator.stop(audioCtx.currentTime + 0.08);
    } catch (e) {
      console.error('Audio beep failed:', e);
    }
  };

  useEffect(() => {
    let html5QrCode = null;
    if (showCameraScanner) {
      const timer = setTimeout(() => {
        try {
          html5QrCode = new Html5Qrcode("camera-reader");
          scannerRef.current = html5QrCode;

          const config = { 
            fps: 10, 
            qrbox: (width, height) => {
              const min = Math.min(width, height);
              const size = Math.floor(min * 0.65);
              return { width: size, height: size };
            }
          };

          let lastScannedBarcode = '';
          let lastScanTime = 0;

          html5QrCode.start(
            { facingMode: "environment" },
            config,
            (decodedText) => {
              const now = Date.now();
              if (decodedText === lastScannedBarcode && now - lastScanTime < 2000) {
                return;
              }
              lastScannedBarcode = decodedText;
              lastScanTime = now;

              playBeep();

              const container = document.getElementById('camera-reader-container');
              if (container) {
                container.style.boxShadow = '0 0 30px #22c55e';
                setTimeout(() => { container.style.boxShadow = 'none'; }, 500);
              }

              const found = agentProducts.find(p => p.barcode === decodedText || p.productBarcode === decodedText);
              if (found) {
                handleAddProductToCart(found);
                showAlert(language === 'uz' ? `Qo'shildi: ${found.productName}` : `Добавлено: ${found.productName}`, 'success');
              } else {
                showAlert(language === 'uz' ? `Shtrix-kod topilmadi: ${decodedText}` : `Штрих-код не найден: ${decodedText}`, 'error');
              }
            },
            () => {}
          ).catch(err => {
            console.error("Camera start error:", err);
            showAlert(language === 'uz' ? "Kamerani ishga tushirib bo'lmadi. Ruxsat berilganini tekshiring." : "Не удалось запустить камеру. Проверьте разрешения.", "error");
          });
        } catch (e) {
          console.error("Scanner init error:", e);
        }
      }, 300);

      return () => {
        clearTimeout(timer);
        if (html5QrCode && html5QrCode.isScanning) {
          html5QrCode.stop().catch(err => console.error("Scanner stop error:", err));
        }
      };
    }
  }, [showCameraScanner, agentProducts, language]);

  const handleAddCheckedProducts = () => {
    if (selectedModalProducts.length === 0) {
      setShowProductModal(false);
      return;
    }
    
    selectedModalProducts.forEach(prodId => {
      const prod = agentProducts.find(p => p.productId === prodId || p.id === prodId);
      if (prod) {
        handleAddProductToCart(prod);
      }
    });
    
    setSelectedModalProducts([]);
    setShowProductModal(false);
    showAlert(language === 'uz' ? "Mahsulotlar muvaffaqiyatli qo'shildi!" : "Товары успешно добавлены!", 'success');
  };


  // Custom Logo Component (Dynamically editable)
  const Logo = () => (
    <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
      {companyLogo ? (
        <img 
          src={companyLogo} 
          style={{ width: '40px', height: '40px', borderRadius: '8px', objectFit: 'contain' }} 
          alt="Logo" 
        />
      ) : (
        <svg width="40" height="40" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
          <polygon points="50,10 90,30 90,70 50,90 10,70 10,30" stroke="var(--accent-color)" strokeWidth="8" strokeLinejoin="round" fill="none" />
          <circle cx="50" cy="50" r="14" fill="var(--accent-color)" />
          <line x1="50" y1="10" x2="50" y2="36" stroke="var(--accent-color)" strokeWidth="6" />
          <line x1="10" y1="70" x2="38" y2="58" stroke="var(--accent-color)" strokeWidth="6" />
          <line x1="90" y1="70" x2="62" y2="58" stroke="var(--accent-color)" strokeWidth="6" />
        </svg>
      )}
      <div>
        <h3 style={{ fontSize: '18px', fontWeight: '800', letterSpacing: '1px', color: companyNameColor }}>{companyName}</h3>
        <span style={{ fontSize: '10px', color: companyBioColor, fontWeight: '600', letterSpacing: '2px', display: 'block', marginTop: '-3px' }}>{companyBio}</span>
      </div>
    </div>
  );
  if (isLoading) {
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        width: '100vw',
        height: '100vh',
        backgroundColor: 'var(--bg-primary)',
        color: 'var(--text-primary)',
        gap: '20px'
      }}>
        <div style={{
          width: '50px',
          height: '50px',
          border: '4px solid rgba(13, 148, 136, 0.1)',
          borderTop: '4px solid var(--accent-color)',
          borderRadius: '50%',
          animation: 'spin 1s linear infinite'
        }} />
        <style>{`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}</style>
        <span style={{ fontSize: '14px', fontWeight: '500', color: 'var(--text-secondary)' }}>
          {language === 'uz' ? 'Bulutli serverga ulanilmoqda...' : 'Подключение к облачному серверу...'}
        </span>
      </div>
    );
  }

  if (!isLoggedIn) {
    return (
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '100vw',
        height: '100vh',
        backgroundColor: 'var(--bg-primary)',
        backgroundImage: 'radial-gradient(circle at top right, rgba(13, 148, 136, 0.08), transparent 40%)'
      }}>
        <div style={{
          width: '400px',
          padding: '40px',
          borderRadius: '16px',
          backgroundColor: 'var(--bg-secondary)',
          border: '1px solid var(--border-color)',
          boxShadow: 'var(--shadow-lg)',
          textAlign: 'center'
        }} className="fade-in">
          
          <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '32px' }}>
            <Logo />
          </div>

          <h2 style={{ fontSize: '22px', marginBottom: '8px', fontWeight: '500' }}>{t('login_title')}</h2>
          <p style={{ color: 'var(--text-secondary)', fontSize: '14px', marginBottom: '24px' }}>{t('login_subtitle')}</p>

          {loginError && (
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              backgroundColor: 'rgba(239, 68, 110, 0.1)',
              border: '1px solid var(--danger-color)',
              color: 'var(--danger-color)',
              padding: '12px',
              borderRadius: '8px',
              fontSize: '14px',
              marginBottom: '20px',
              textAlign: 'left'
            }}>
              <AlertCircle size={16} />
              <span>{loginError.includes('xato') || loginError.includes('Xato') ? t('login_error') : loginError}</span>
            </div>
          )}

          <form onSubmit={handleLogin} style={{ display: 'flex', flexDirection: 'column', gap: '16px', textAlign: 'left' }}>
            <div>
              <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('username')}</label>
              <input 
                type="text" 
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="admin / ombor"
                style={{
                  width: '100%',
                  padding: '12px 16px',
                  borderRadius: '8px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: 'var(--bg-primary)',
                  color: 'var(--text-primary)',
                  fontSize: '14px'
                }}
              />
            </div>

            <div>
              <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('password')}</label>
              <input 
                type="password" 
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="••••••••"
                style={{
                  width: '100%',
                  padding: '12px 16px',
                  borderRadius: '8px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: 'var(--bg-primary)',
                  color: 'var(--text-primary)',
                  fontSize: '14px'
                }}
              />
            </div>

            <button 
              type="submit"
              style={{
                width: '100%',
                padding: '12px',
                borderRadius: '8px',
                border: 'none',
                backgroundColor: 'var(--accent-color)',
                color: '#fff',
                fontSize: '15px',
                fontWeight: '600',
                cursor: 'pointer',
                marginTop: '8px',
                transition: 'background var(--transition-fast)'
              }}
              onMouseOver={(e) => e.target.style.backgroundColor = 'var(--accent-hover)'}
              onMouseOut={(e) => e.target.style.backgroundColor = 'var(--accent-color)'}
            >
              {t('login_btn')}
            </button>
          </form>

          <div style={{ marginTop: '24px', fontSize: '12px', color: 'var(--text-muted)' }}>
            {t('demo_mode')}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', width: '100%', height: '100%', backgroundColor: 'var(--bg-primary)' }}>
      {/* Sidebar Overlay for mobile */}
      {isMobileMenuOpen && (
        <div className="sidebar-overlay" onClick={() => setIsMobileMenuOpen(false)} />
      )}

      {/* Sidebar Panel */}
      <aside 
        className={`sidebar ${isMobileMenuOpen ? 'open' : ''}`}
        style={{
          width: '260px',
          backgroundColor: 'var(--bg-secondary)',
          borderRight: '1px solid var(--border-color)',
          display: 'flex',
          flexDirection: 'column',
          padding: '24px 16px 48px 16px', // Increased bottom padding for mobile safety
          overflowY: 'auto'
        }}
      >
        <div style={{ marginBottom: '40px', paddingLeft: '8px' }}>
          <Logo />
        </div>

        {/* Menu List */}
        <nav style={{ display: 'flex', flexDirection: 'column', gap: '8px', flexGrow: 1 }}>
          {userRole !== 'agent' && (
            <button 
              onClick={() => setActiveTab('dashboard')}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '12px',
                width: '100%',
                padding: '12px 16px',
                borderRadius: '8px',
                border: 'none',
                backgroundColor: activeTab === 'dashboard' ? 'var(--accent-light)' : 'transparent',
                color: activeTab === 'dashboard' ? 'var(--accent-color)' : 'var(--text-secondary)',
                cursor: 'pointer',
                fontWeight: '500',
                textAlign: 'left',
                transition: 'all var(--transition-fast)'
              }}
            >
              <LayoutDashboard size={18} />
              <span>{t('dashboard')}</span>
            </button>
          )}

          {userRole !== 'agent' && (
            <button 
              onClick={() => setActiveTab('products')}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '12px',
                width: '100%',
                padding: '12px 16px',
                borderRadius: '8px',
                border: 'none',
                backgroundColor: activeTab === 'products' ? 'var(--accent-light)' : 'transparent',
                color: activeTab === 'products' ? 'var(--accent-color)' : 'var(--text-secondary)',
                cursor: 'pointer',
                fontWeight: '500',
                textAlign: 'left',
                transition: 'all var(--transition-fast)'
              }}
            >
              <Package size={18} />
              <span>{t('products')}</span>
            </button>
          )}

          {userRole === 'admin' && (
            <button 
              onClick={() => setActiveTab('stores')}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '12px',
                width: '100%',
                padding: '12px 16px',
                borderRadius: '8px',
                border: 'none',
                backgroundColor: activeTab === 'stores' ? 'var(--accent-light)' : 'transparent',
                color: activeTab === 'stores' ? 'var(--accent-color)' : 'var(--text-secondary)',
                cursor: 'pointer',
                fontWeight: '500',
                textAlign: 'left',
                transition: 'all var(--transition-fast)'
              }}
            >
              <StoreIcon size={18} />
              <span>{t('stores')}</span>
            </button>
          )}


          {userRole === 'admin' && (
            <div style={{ display: 'flex', flexDirection: 'column' }}>
              <button 
                onClick={() => setTahlillarOpen(!tahlillarOpen)}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  width: '100%',
                  padding: '12px 16px',
                  borderRadius: '8px',
                  border: 'none',
                  backgroundColor: ['tahlil_umumiy', 'tahlil_dokon', 'tahlil_agent'].includes(activeTab) ? 'var(--accent-light)' : 'transparent',
                  color: ['tahlil_umumiy', 'tahlil_dokon', 'tahlil_agent'].includes(activeTab) ? 'var(--accent-color)' : 'var(--text-secondary)',
                  cursor: 'pointer',
                  fontWeight: '500',
                  textAlign: 'left',
                  transition: 'all var(--transition-fast)'
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <TrendingUp size={18} />
                  <span>{t('analytics')}</span>
                </div>
                <ChevronDown 
                  size={16} 
                  style={{ 
                    transform: tahlillarOpen ? 'rotate(180deg)' : 'rotate(0deg)', 
                    transition: 'transform 0.2s',
                    color: 'var(--text-muted)'
                  }} 
                />
              </button>
              
              {tahlillarOpen && (
                <div style={{ 
                  display: 'flex', 
                  flexDirection: 'column', 
                  gap: '4px', 
                  paddingLeft: '28px', 
                  marginTop: '4px'
                }}>
                  <button
                    onClick={() => setActiveTab('tahlil_umumiy')}
                    style={{
                      padding: '8px 12px',
                      borderRadius: '6px',
                      border: 'none',
                      backgroundColor: activeTab === 'tahlil_umumiy' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                      color: activeTab === 'tahlil_umumiy' ? 'var(--accent-color)' : 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '13px',
                      fontWeight: '500',
                      textAlign: 'left',
                      transition: 'all var(--transition-fast)'
                    }}
                  >
                    {t('general_analytics')}
                  </button>
                  <button
                    onClick={() => setActiveTab('tahlil_dokon')}
                    style={{
                      padding: '8px 12px',
                      borderRadius: '6px',
                      border: 'none',
                      backgroundColor: activeTab === 'tahlil_dokon' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                      color: activeTab === 'tahlil_dokon' ? 'var(--accent-color)' : 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '13px',
                      fontWeight: '500',
                      textAlign: 'left',
                      transition: 'all var(--transition-fast)'
                    }}
                  >
                    {t('store_analytics')}
                  </button>
                  <button
                    onClick={() => setActiveTab('tahlil_agent')}
                    style={{
                      padding: '8px 12px',
                      borderRadius: '6px',
                      border: 'none',
                      backgroundColor: activeTab === 'tahlil_agent' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                      color: activeTab === 'tahlil_agent' ? 'var(--accent-color)' : 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '13px',
                      fontWeight: '500',
                      textAlign: 'left',
                      transition: 'all var(--transition-fast)'
                    }}
                  >
                    {t('agent_analytics')}
                  </button>
                </div>
              )}
            </div>
          )}

          <button 
            onClick={() => setActiveTab('assignments')}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '12px',
              width: '100%',
              padding: '12px 16px',
              borderRadius: '8px',
              border: 'none',
              backgroundColor: activeTab === 'assignments' ? 'var(--accent-light)' : 'transparent',
              color: activeTab === 'assignments' ? 'var(--accent-color)' : 'var(--text-secondary)',
              cursor: 'pointer',
              fontWeight: '500',
              textAlign: 'left',
              transition: 'all var(--transition-fast)'
            }}
          >
            <Users size={18} />
            <span>{userRole === 'agent' ? t('my_tasks') : t('assignments')}</span>
          </button>

          {userRole === 'agent' && (
            <button 
              onClick={() => setActiveTab('agent_history')}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '12px',
                width: '100%',
                padding: '12px 16px',
                borderRadius: '8px',
                border: 'none',
                backgroundColor: activeTab === 'agent_history' ? 'var(--accent-light)' : 'transparent',
                color: activeTab === 'agent_history' ? 'var(--accent-color)' : 'var(--text-secondary)',
                cursor: 'pointer',
                fontWeight: '500',
                textAlign: 'left',
                transition: 'all var(--transition-fast)'
              }}
            >
              <Clock size={18} />
              <span>{language === 'uz' ? "Tarix" : "История"}</span>
            </button>
          )}

          {userRole !== 'agent' && (
            <button 
              onClick={() => setActiveTab('sales')}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '12px',
                width: '100%',
                padding: '12px 16px',
                borderRadius: '8px',
                border: 'none',
                backgroundColor: activeTab === 'sales' ? 'var(--accent-light)' : 'transparent',
                color: activeTab === 'sales' ? 'var(--accent-color)' : 'var(--text-secondary)',
                cursor: 'pointer',
                fontWeight: '500',
                textAlign: 'left',
                transition: 'all var(--transition-fast)'
              }}
            >
              <FileText size={18} />
              <span>{t('sales_history')}</span>
            </button>
          )}

          {['admin', 'agent'].includes(userRole) && (
            <div style={{ display: 'flex', flexDirection: 'column' }}>
              <button 
                onClick={() => setSettingsOpen(!settingsOpen)}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  width: '100%',
                  padding: '12px 16px',
                  borderRadius: '8px',
                  border: 'none',
                  backgroundColor: ['settings_profile', 'settings_language', 'settings_discounts', 'settings_payments', 'settings_company', 'settings_agents', 'settings_admins'].includes(activeTab) ? 'var(--accent-light)' : 'transparent',
                  color: ['settings_profile', 'settings_language', 'settings_discounts', 'settings_payments', 'settings_company', 'settings_agents', 'settings_admins'].includes(activeTab) ? 'var(--accent-color)' : 'var(--text-secondary)',
                  cursor: 'pointer',
                  fontWeight: '500',
                  textAlign: 'left',
                  transition: 'all var(--transition-fast)'
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <Settings size={18} />
                  <span>{t('settings')}</span>
                </div>
                <ChevronDown 
                  size={16} 
                  style={{ 
                    transform: settingsOpen ? 'rotate(180deg)' : 'rotate(0deg)', 
                    transition: 'transform 0.2s',
                    color: 'var(--text-muted)'
                  }} 
                />
              </button>
              
              {settingsOpen && (
                <div style={{ 
                  display: 'flex', 
                  flexDirection: 'column', 
                  gap: '4px', 
                  paddingLeft: '28px', 
                  marginTop: '4px'
                }}>
                  {userRole === 'admin' && (
                    <button
                      onClick={() => setActiveTab('settings_profile')}
                      style={{
                        padding: '8px 12px',
                        borderRadius: '6px',
                        border: 'none',
                        backgroundColor: activeTab === 'settings_profile' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                        color: activeTab === 'settings_profile' ? 'var(--accent-color)' : 'var(--text-secondary)',
                        cursor: 'pointer',
                        fontSize: '13px',
                        fontWeight: '500',
                        textAlign: 'left',
                        transition: 'all var(--transition-fast)'
                      }}
                    >
                      {t('profile_settings')}
                    </button>
                  )}
                  <button
                    onClick={() => setActiveTab('settings_language')}
                    style={{
                      padding: '8px 12px',
                      borderRadius: '6px',
                      border: 'none',
                      backgroundColor: activeTab === 'settings_language' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                      color: activeTab === 'settings_language' ? 'var(--accent-color)' : 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '13px',
                      fontWeight: '500',
                      textAlign: 'left',
                      transition: 'all var(--transition-fast)'
                    }}
                  >
                    {t('language_settings')}
                  </button>
{userRole === 'admin' && (
                    <button
                      onClick={() => setActiveTab('settings_discounts')}
                      style={{
                        padding: '8px 12px',
                        borderRadius: '6px',
                        border: 'none',
                        backgroundColor: activeTab === 'settings_discounts' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                        color: activeTab === 'settings_discounts' ? 'var(--accent-color)' : 'var(--text-secondary)',
                        cursor: 'pointer',
                        fontSize: '13px',
                        fontWeight: '500',
                        textAlign: 'left',
                        transition: 'all var(--transition-fast)'
                      }}
                    >
                      {t('settings_discounts')}
                    </button>
                  )}
                  <button
                    onClick={() => setActiveTab('settings_payments')}
                    style={{
                      padding: '8px 12px',
                      borderRadius: '6px',
                      border: 'none',
                      backgroundColor: activeTab === 'settings_payments' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                      color: activeTab === 'settings_payments' ? 'var(--accent-color)' : 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '13px',
                      fontWeight: '500',
                      textAlign: 'left',
                      transition: 'all var(--transition-fast)'
                    }}
                  >
                    {language === 'uz' ? "Terminal Sozlamalari" : 'Настройки терминала'}
                  </button>
                  {userRole === 'admin' && (
                    <>
                      <button
                        onClick={() => setActiveTab('settings_company')}
                        style={{
                          padding: '8px 12px',
                          borderRadius: '6px',
                          border: 'none',
                          backgroundColor: activeTab === 'settings_company' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                          color: activeTab === 'settings_company' ? 'var(--accent-color)' : 'var(--text-secondary)',
                          cursor: 'pointer',
                          fontSize: '13px',
                          fontWeight: '500',
                          textAlign: 'left',
                          transition: 'all var(--transition-fast)'
                        }}
                      >
                        {language === 'uz' ? 'Firma Sozlamalari' : 'Настройки фирмы'}
                      </button>
                      <button
                        onClick={() => setActiveTab('settings_agents')}
                        style={{
                          padding: '8px 12px',
                          borderRadius: '6px',
                          border: 'none',
                          backgroundColor: activeTab === 'settings_agents' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                          color: activeTab === 'settings_agents' ? 'var(--accent-color)' : 'var(--text-secondary)',
                          cursor: 'pointer',
                          fontSize: '13px',
                          fontWeight: '500',
                          textAlign: 'left',
                          transition: 'all var(--transition-fast)'
                        }}
                      >
                        {language === 'uz' ? 'Agentlar Boshqaruvi' : 'Управление агентами'}
                      </button>
                      <button
                        onClick={() => setActiveTab('settings_admins')}
                        style={{
                          padding: '8px 12px',
                          borderRadius: '6px',
                          border: 'none',
                          backgroundColor: activeTab === 'settings_admins' ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                          color: activeTab === 'settings_admins' ? 'var(--accent-color)' : 'var(--text-secondary)',
                          cursor: 'pointer',
                          fontSize: '13px',
                          fontWeight: '500',
                          textAlign: 'left',
                          transition: 'all var(--transition-fast)'
                        }}
                      >
                        {language === 'uz' ? 'Adminlar Boshqaruvi' : 'Управление админами'}
                      </button>
                    </>
                  )}
                  
                  {/* Compact Logout button inside Settings submenu */}
                  <button
                    onClick={handleLogout}
                    style={{
                      padding: '6px 12px',
                      borderRadius: '6px',
                      border: 'none',
                      backgroundColor: 'transparent',
                      color: 'var(--danger-color)',
                      cursor: 'pointer',
                      fontSize: '12px',
                      fontWeight: '600',
                      textAlign: 'left',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px',
                      marginTop: '6px',
                      borderTop: '1px dashed rgba(239, 68, 68, 0.15)',
                      paddingTop: '8px',
                      transition: 'all var(--transition-fast)'
                    }}
                  >
                    <LogOut size={14} />
                    <span>{t('logout')}</span>
                  </button>
                </div>
              )}
            </div>
          )}
          

        </nav>

        {/* User Info & Logout */}
        <div style={{
          borderTop: '1px solid var(--border-color)',
          paddingTop: '20px',
          display: 'flex',
          flexDirection: 'column',
          gap: '12px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{
              width: '36px',
              height: '36px',
              borderRadius: '50%',
              backgroundColor: 'var(--bg-tertiary)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              fontWeight: '600',
              fontSize: '14px',
              overflow: 'hidden'
            }}>
              {userRole === 'admin' 
                ? (adminPhoto 
                    ? <img src={adminPhoto} style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover' }} alt="Admin" />
                    : 'AD'
                  )
                : userRole === 'warehouse_manager' ? 'OB' : 'AG'}
            </div>
            <div>
              <div style={{ fontSize: '13px', fontWeight: '600' }}>
                {userRole === 'admin' 
                  ? adminName 
                  : userRole === 'warehouse_manager' 
                    ? (language === 'uz' ? 'Ombor Mudiri' : 'Зав. складом') 
                    : (language === 'uz' ? 'Agent' : 'Агент')}
              </div>
              <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{t('active_in_system')}</span>
            </div>
          </div>


        </div>
      </aside>

      {/* Main Panel Content */}
      <main style={{
        flexGrow: 1,
        display: 'flex',
        flexDirection: 'column',
        overflow: 'hidden'
      }}>
        {/* Top Header */}
        <header 
          className="top-header"
          style={{
            height: '70px',
            borderBottom: '1px solid var(--border-color)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '0 32px',
            backgroundColor: 'var(--bg-secondary)'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <button 
              className="mobile-menu-btn" 
              onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
              style={{ 
                marginRight: '12px',
                background: 'transparent',
                border: 'none',
                color: 'var(--text-primary)',
                cursor: 'pointer',
                display: 'none', // Overridden in CSS media queries to show on mobile
                alignItems: 'center',
                justifyContent: 'center',
                padding: '4px'
              }}
            >
              <Menu size={22} />
            </button>
            <h2 style={{ fontSize: '20px', fontWeight: '500' }}>
            {activeTab === 'dashboard' && t('dashboard')}
            {activeTab === 'products' && t('products_catalog')}
            {activeTab === 'stores' && t('stores_list')}
            {(activeTab === 'agents' || activeTab === 'settings_agents') && (language === 'uz' ? 'Agentlar Boshqaruvi' : 'Управление агентами')}
            {activeTab === 'settings_admins' && (language === 'uz' ? 'Adminlar Boshqaruvi' : 'Управление администраторами')}
            {activeTab === 'assignments' && (userRole === 'agent' ? t('my_tasks') : t('agent_assignments'))}
            {activeTab === 'agent_history' && (language === 'uz' ? 'Tarix' : 'История')}
            {activeTab === 'sales' && t('sales_history_title')}
            {activeTab === 'tahlil_umumiy' && t('general_analytics')}
            {activeTab === 'tahlil_dokon' && t('store_analytics')}
            {activeTab === 'tahlil_agent' && t('agent_analytics')}
            {activeTab === 'settings_profile' && t('profile_title')}
            {activeTab === 'settings_language' && t('language_title')}
            {activeTab === 'settings_discounts' && t('settings_discounts')}
          </h2>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
            <button
              onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: '36px',
                height: '36px',
                borderRadius: '8px',
                backgroundColor: 'var(--bg-primary)',
                border: '1px solid var(--border-color)',
                color: 'var(--text-primary)',
                cursor: 'pointer',
                transition: 'all var(--transition-fast)'
              }}
              title={theme === 'dark' ? (language === 'uz' ? 'Kunduzgi rejim' : 'Светлая тема') : (language === 'uz' ? 'Tungi rejim' : 'Темная тема')}
            >
              {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
            </button>
            <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
              {language === 'uz' ? 'Sana' : 'Дата'}: {new Date().toLocaleDateString(language === 'uz' ? 'uz-UZ' : 'ru-RU')}
            </span>
            <span style={{
              fontSize: '11px',
              fontWeight: '600',
              backgroundColor: 'rgba(13, 148, 136, 0.15)',
              color: 'var(--accent-color)',
              padding: '4px 10px',
              borderRadius: '12px',
              textTransform: 'uppercase'
            }}>
              {userRole === 'admin' ? (language === 'uz' ? 'Admin' : 'Админ') : userRole === 'warehouse_manager' ? (language === 'uz' ? 'Omborchi' : 'Кладовщик') : (language === 'uz' ? 'Agent' : 'Агент')}
            </span>
          </div>
        </header>

        {/* Content Body scrollable */}
        <div className="content-body" style={{ flexGrow: 1, overflowY: 'auto', padding: '32px' }}>
          
          {/* VIEW 1: DASHBOARD */}
          {activeTab === 'dashboard' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '32px' }}>
              
              {/* Date Filter selector */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap', alignSelf: 'flex-start' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '8px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <span style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: '500' }}>{t('date_filter')}:</span>
                  <select 
                    value={dateFilter} 
                    onChange={(e) => setDateFilter(e.target.value)}
                    style={{
                      backgroundColor: 'transparent',
                      border: 'none',
                      color: 'var(--accent-color)',
                      fontSize: '13px',
                      fontWeight: '600',
                      cursor: 'pointer',
                      outline: 'none'
                    }}
                  >
                    <option value="bugun">{t('bugun')}</option>
                    <option value="custom">{t('custom_range')}</option>
                  </select>
                </div>

                {dateFilter === 'custom' && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '6px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', animation: 'fadeIn 0.2s ease' }}>
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('dan')}</span>
                    <input 
                      type="date" 
                      value={startDate} 
                      onChange={(e) => setStartDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('gacha')}</span>
                    <input 
                      type="date" 
                      value={endDate} 
                      onChange={(e) => setEndDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                  </div>
                )}
              </div>

              {(() => {
                const filteredSales = getFilteredSales(sales);
                const totalRevenue = filteredSales.reduce((sum, s) => sum + s.total, 0);
                
                let totalSoldQty = 0;
                let totalCost = 0;
                const activeAgentsSet = new Set();
                filteredSales.forEach(s => {
                  activeAgentsSet.add(s.agent);
                  if (s.items) {
                    s.items.forEach(item => {
                      totalSoldQty += item.qty;
                      const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
                      totalCost += (cost * item.qty);
                    });
                  }
                });
                const totalProfit = totalRevenue - totalCost;

                // Prepare chart data dynamically
                const dailyData = {};
                filteredSales.forEach(s => {
                  const date = s.date;
                  if (!dailyData[date]) dailyData[date] = { name: date, savdo: 0 };
                  dailyData[date].savdo += s.total;
                });
                const dynamicChartData = Object.keys(dailyData).length > 0 
                  ? Object.values(dailyData).sort((a, b) => new Date(a.name) - new Date(b.name))
                  : [{ name: 'Ma\'lumot yo\'q', savdo: 0 }];

                return (
                  <>
                    {/* Stats Cards Grid */}
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '24px' }}>
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '8px', backgroundColor: 'var(--accent-light)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent-color)' }}>
                          <DollarSign size={24} style={{ alignSelf: 'center' }} />
                        </div>
                        <div>
                          <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '4px' }}>{t('revenue')}</div>
                          <div style={{ fontSize: '20px', fontWeight: '700' }}>{totalRevenue.toLocaleString()} UZS</div>
                        </div>
                      </div>

                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '8px', backgroundColor: 'rgba(16, 185, 129, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--success-color)' }}>
                          <UserCheck size={24} style={{ alignSelf: 'center' }} />
                        </div>
                        <div>
                          <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '4px' }}>{t('active_agents')}</div>
                          <div style={{ fontSize: '20px', fontWeight: '700' }}>{activeAgentsSet.size} {language === 'uz' ? 'nafar' : 'чел.'}</div>
                        </div>
                      </div>

                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '8px', backgroundColor: 'rgba(245, 158, 11, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--warning-color)' }}>
                          <ShoppingCart size={24} style={{ alignSelf: 'center' }} />
                        </div>
                        <div>
                          <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '4px' }}>{t('sold_items')}</div>
                          <div style={{ fontSize: '20px', fontWeight: '700' }}>{totalSoldQty.toLocaleString()} {t('dona')}</div>
                        </div>
                      </div>

                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '8px', backgroundColor: 'rgba(99, 102, 241, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#6366f1' }}>
                          <TrendingUp size={24} style={{ alignSelf: 'center' }} />
                        </div>
                        <div>
                          <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '4px' }}>{t('net_profit')}</div>
                          <div style={{ fontSize: '20px', fontWeight: '700' }}>{totalProfit.toLocaleString()} UZS</div>
                        </div>
                      </div>
                    </div>

                    {/* Chart & Activity Block */}
                    <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '24px' }}>
                      {/* Sales Chart */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                        <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '20px' }}>{language === 'uz' ? 'Sotuvlar dinamikasi (UZS)' : 'Динамика продаж (UZS)'}</h3>
                        <div style={{ width: '100%', height: '300px' }}>
                          <ResponsiveContainer width="100%" height="100%">
                            <AreaChart data={dynamicChartData} margin={{ top: 10, right: 10, left: 10, bottom: 0 }}>
                              <defs>
                                <linearGradient id="savdoColor" x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor="var(--accent-color)" stopOpacity={0.4}/>
                                  <stop offset="95%" stopColor="var(--accent-color)" stopOpacity={0.0}/>
                                </linearGradient>
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" stroke="var(--border-color)" />
                              <XAxis dataKey="name" stroke="var(--text-muted)" fontSize={12} />
                              <YAxis stroke="var(--text-muted)" fontSize={12} tickFormatter={(v) => `${(v || 0)/1000}k`} />
                              <Tooltip contentStyle={{ backgroundColor: 'var(--bg-secondary)', borderColor: 'var(--border-color)', color: 'var(--text-primary)' }} />
                              <Area type="monotone" dataKey="savdo" stroke="var(--accent-color)" strokeWidth={3} fillOpacity={1} fill="url(#savdoColor)" />
                            </AreaChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                      {/* Live Agent Activity List */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column' }}>
                        <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '20px' }}>{t('active_agents_today')}</h3>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', overflowY: 'auto', flexGrow: 1 }}>
                          {agents.map((agent) => (
                            <div key={agent.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingBottom: '12px', borderBottom: '1px solid var(--border-color)' }}>
                              <div>
                                <div style={{ fontSize: '14px', fontWeight: '600' }}>{agent.login}</div>
                                <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>@{agent.username || 'username'}</span>
                              </div>
                              <span style={{ fontSize: '12px', color: 'var(--success-color)', display: 'flex', alignItems: 'center', gap: '4px' }}>
                                <span style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: 'var(--success-color)' }}></span>
                                {language === 'uz' ? 'online' : 'в сети'}
                              </span>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>

                    {/* Recent Sales Table */}
                    <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                      <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '16px' }}>{language === 'uz' ? 'Oxirgi toʻlovlar' : 'Последние платежи'}</h3>
                      <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                        <thead>
                          <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                            <th style={{ padding: '12px' }}>{language === 'uz' ? 'Tranzaksiya ID' : 'ID Транзакции'}</th>
                            <th style={{ padding: '12px' }}>{language === 'uz' ? 'Agent' : 'Агент'}</th>
                            <th style={{ padding: '12px' }}>{language === 'uz' ? 'Do\'kon nomi' : 'Название магазина'}</th>
                            <th style={{ padding: '12px' }}>{language === 'uz' ? 'To\'lov summasi' : 'Сумма оплаты'}</th>
                            <th style={{ padding: '12px' }}>{language === 'uz' ? 'To\'lov turi' : 'Тип оплаты'}</th>
                            <th style={{ padding: '12px' }}>{language === 'uz' ? 'Holat' : 'Статус'}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredSales.map((sale) => (
                            <tr key={sale.id} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}>
                              <td style={{ padding: '12px', color: 'var(--accent-color)', fontWeight: '600' }}>#TR-{sale.id}</td>
                              <td style={{ padding: '12px' }}>{getAgentLoginByName(sale.agent)}</td>
                              <td style={{ padding: '12px' }}>{sale.store}</td>
                              <td style={{ padding: '12px', fontWeight: '600' }}>{sale.total.toLocaleString()} UZS</td>
                              <td style={{ padding: '12px' }}>
                                <span style={{
                                  fontSize: '11px',
                                  fontWeight: '600',
                                  backgroundColor: sale.payment === 'Payme' ? 'rgba(13, 148, 136, 0.15)' : 'rgba(99, 102, 241, 0.15)',
                                  color: sale.payment === 'Payme' ? 'var(--accent-color)' : '#6366f1',
                                  padding: '3px 8px',
                                  borderRadius: '4px'
                                }}>{sale.payment}</span>
                              </td>
                              <td style={{ padding: '12px' }}>
                                <span style={{
                                  fontSize: '11px',
                                  fontWeight: '600',
                                  backgroundColor: 'rgba(16, 185, 129, 0.15)',
                                  color: 'var(--success-color)',
                                  padding: '3px 8px',
                                  borderRadius: '4px',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px'
                                }}>
                                  <Check size={12} /> {language === 'uz' ? 'to\'landi' : 'оплачено'}
                                </span>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                );
              })()}
            </div>
          )}

          {/* VIEW 2: PRODUCTS */}
          {activeTab === 'products' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              
              {/* Tab Toggle for Products Sub-views */}
              <div style={{ display: 'flex', gap: '8px', borderBottom: '1px solid var(--border-color)', paddingBottom: '12px' }}>
                <button
                  onClick={() => setProductsTabMode('warehouse')}
                  style={{
                    padding: '10px 20px',
                    borderRadius: '8px',
                    border: 'none',
                    backgroundColor: productsTabMode === 'warehouse' ? 'var(--accent-color)' : 'transparent',
                    color: productsTabMode === 'warehouse' ? '#fff' : 'var(--text-secondary)',
                    fontWeight: '600',
                    fontSize: '14px',
                    cursor: 'pointer',
                    transition: 'all var(--transition-fast)'
                  }}
                >
                  {t('warehouse_stock')}
                </button>
                <button
                  id="hozirda-mavjud-btn"
                  onClick={() => setProductsTabMode('agents_stock')}
                  style={{
                    padding: '10px 20px',
                    borderRadius: '8px',
                    border: 'none',
                    backgroundColor: productsTabMode === 'agents_stock' ? 'var(--accent-color)' : 'transparent',
                    color: productsTabMode === 'agents_stock' ? '#fff' : 'var(--text-secondary)',
                    fontWeight: '600',
                    fontSize: '14px',
                    cursor: 'pointer',
                    transition: 'all var(--transition-fast)'
                  }}
                >
                  {t('current_agent_stock')}
                </button>
              </div>

              {productsTabMode === 'warehouse' ? (
                <>
                  {/* Product Tools Header */}
                  <div style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    flexWrap: 'wrap',
                    gap: '16px',
                    backgroundColor: 'var(--bg-secondary)',
                    padding: '20px 24px',
                    borderRadius: '12px',
                    border: '1px solid var(--border-color)'
                  }}>
                    {/* Search Bar */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexGrow: 1, maxWidth: '400px', position: 'relative' }}>
                      <Search size={18} style={{ position: 'absolute', left: '12px', color: 'var(--text-muted)' }} />
                      <input 
                        type="text" 
                        value={productSearchQuery}
                        onChange={(e) => setProductSearchQuery(e.target.value)}
                        placeholder={language === 'uz' ? "Nomi yoki shtrix kod bo'yicha qidirish..." : "Поиск по названию или штрих-коду..."}
                        style={{
                          width: '100%',
                          padding: '10px 12px 10px 40px',
                          borderRadius: '8px',
                          border: '1px solid var(--border-color)',
                          backgroundColor: 'var(--bg-primary)',
                          color: 'var(--text-primary)',
                          fontSize: '14px'
                        }}
                      />
                    </div>

                    {/* Barcode/QR Scanner simulation input */}
                    <form onSubmit={handleScannerSubmit} style={{ display: 'flex', alignItems: 'center', gap: '10px', flexGrow: 1, maxWidth: '380px', position: 'relative' }}>
                      <div style={{ position: 'absolute', left: '12px', color: 'var(--accent-color)', display: 'flex', alignItems: 'center' }}>
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 5H5V19H3V5ZM21 5H19V19H21V5ZM7 5H8V19H7V5ZM11 5H13V19H11V5ZM15 5H17V19H15V5ZM9 5H10V19H9V5Z"/></svg>
                      </div>
                      <input 
                        type="text" 
                        value={scannerValue}
                        onChange={(e) => setScannerValue(e.target.value)}
                        placeholder={language === 'uz' ? "Shtrix-kod / QR skanerlash maydoni..." : "Поле сканирования штрих-кода / QR..."}
                        style={{
                          width: '100%',
                          padding: '10px 12px 10px 40px',
                          borderRadius: '8px',
                          border: '1px solid var(--accent-border)',
                          backgroundColor: 'var(--accent-light)',
                          color: 'var(--text-primary)',
                          fontSize: '14px'
                        }}
                      />
                      <button 
                        type="submit" 
                        style={{
                          padding: '10px 16px',
                          borderRadius: '8px',
                          border: 'none',
                          backgroundColor: 'var(--accent-color)',
                          color: '#fff',
                          fontSize: '13px',
                          fontWeight: '600',
                          cursor: 'pointer'
                        }}
                      >
                        {language === 'uz' ? 'Skanerlash' : 'Сканировать'}
                      </button>
                    </form>

                    {/* Add Product Button */}
                    <button 
                      onClick={() => {
                        setNewProduct({ barcode: '', name: '', price: '', unit: 'dona', stock: '' });
                        setShowAddProductModal(true);
                      }}
                      style={{
                        padding: '10px 20px',
                        borderRadius: '8px',
                        border: 'none',
                        backgroundColor: 'var(--accent-color)',
                        color: '#fff',
                        fontSize: '14px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px'
                      }}
                    >
                      <Plus size={16} /> {language === 'uz' ? 'Yangi mahsulot' : 'Новый продукт'}
                    </button>
                  </div>

                  {/* Notification Banner for scan status */}
                  {scannerNotification && (
                    <div style={{
                      padding: '12px 20px',
                      borderRadius: '8px',
                      backgroundColor: 'var(--bg-secondary)',
                      borderLeft: '4px solid var(--accent-color)',
                      color: 'var(--text-primary)',
                      fontSize: '14px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '12px'
                    }} className="fade-in">
                      <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: 'var(--accent-color)' }}></div>
                      <span>{scannerNotification}</span>
                    </div>
                  )}

                  {/* Products Table */}
                  <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                      <h3 style={{ fontSize: '16px', fontWeight: '500' }}>{language === 'uz' ? 'Ombordagi barcha tovarlar' : 'Все товары на складе'}</h3>
                      <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{t('all_products_count', { count: products.length })}</span>
                    </div>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                      <thead>
                        <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                          <th style={{ padding: '12px' }}>{t('barcode')}</th>
                          <th style={{ padding: '12px' }}>{t('product_name')}</th>
                          <th style={{ padding: '12px' }}>{t('selling_price')}</th>
                          <th style={{ padding: '12px' }}>{t('original_price')}</th>
                          <th style={{ padding: '12px' }}>{t('unit')}</th>
                          <th style={{ padding: '12px', textAlign: 'right' }}>{t('stock_qty')}</th>
                          <th style={{ padding: '12px', textAlign: 'right' }}>{language === 'uz' ? 'O\'chirish' : 'Удалить'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {products
                          .filter(p => p.name.toLowerCase().includes(productSearchQuery.toLowerCase()) || p.barcode.includes(productSearchQuery))
                          .map((product) => (
                            <tr key={product.id} style={{ 
                              borderBottom: '1px solid var(--border-color)', 
                              fontSize: '14px',
                              backgroundColor: productSearchQuery && product.barcode === productSearchQuery ? 'rgba(13, 148, 136, 0.1)' : 'transparent',
                              transition: 'background-color 0.3s'
                            }}>
                              <td style={{ padding: '12px', fontFamily: 'monospace', color: 'var(--text-secondary)' }}>{product.barcode}</td>
                              <td 
                                onClick={() => setSelectedTahlilProductDetails(product.name)}
                                className="hoverable-row"
                                style={{ 
                                  padding: '12px', 
                                  fontWeight: '500', 
                                  cursor: 'pointer', 
                                  color: 'var(--accent-color)',
                                  textDecoration: 'underline',
                                  textDecorationStyle: 'dashed'
                                }}
                              >
                                {product.name}
                              </td>
                              <td style={{ padding: '12px', fontWeight: '600' }}>{product.price.toLocaleString()} UZS</td>
                              <td style={{ padding: '12px', fontWeight: '500', color: 'var(--text-secondary)' }}>
                                {product.originalPrice ? `${product.originalPrice.toLocaleString()} UZS` : <span style={{ color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '12px' }}>{t('not_entered')}</span>}
                              </td>
                              <td style={{ padding: '12px' }}>
                                <span style={{ backgroundColor: 'var(--bg-tertiary)', padding: '3px 8px', borderRadius: '4px', fontSize: '12px' }}>
                                  {product.unit}
                                </span>
                              </td>
                              <td style={{ padding: '12px', textAlign: 'right' }}>
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: '6px' }}>
                                  <button 
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      const currentStock = parseInt(product.stock) || 0;
                                      if (currentStock > 0) {
                                        handleUpdateProductStock(product.id, currentStock - 1);
                                      }
                                    }}
                                    style={{
                                      padding: '2px 8px',
                                      borderRadius: '4px',
                                      border: '1px solid var(--border-color)',
                                      backgroundColor: 'var(--bg-primary)',
                                      color: 'var(--text-primary)',
                                      cursor: 'pointer',
                                      fontWeight: 'bold',
                                      fontSize: '14px',
                                      display: 'flex',
                                      alignItems: 'center',
                                      justifyContent: 'center',
                                      width: '24px',
                                      height: '24px',
                                      userSelect: 'none'
                                    }}
                                  >
                                    -
                                  </button>
                                  <input 
                                    type="number"
                                    min="0"
                                    value={product.stock}
                                    onChange={(e) => handleUpdateProductStock(product.id, e.target.value)}
                                    onClick={(e) => e.stopPropagation()}
                                    style={{
                                      width: '70px',
                                      padding: '4px 8px',
                                      borderRadius: '6px',
                                      border: '1px solid var(--border-color)',
                                      backgroundColor: 'var(--bg-primary)',
                                      color: 'var(--text-primary)',
                                      textAlign: 'center',
                                      fontWeight: '700',
                                      fontSize: '14px',
                                      outline: 'none'
                                    }}
                                  />
                                  <button 
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      const currentStock = parseInt(product.stock) || 0;
                                      handleUpdateProductStock(product.id, currentStock + 1);
                                    }}
                                    style={{
                                      padding: '2px 8px',
                                      borderRadius: '4px',
                                      border: '1px solid var(--border-color)',
                                      backgroundColor: 'var(--bg-primary)',
                                      color: 'var(--text-primary)',
                                      cursor: 'pointer',
                                      fontWeight: 'bold',
                                      fontSize: '14px',
                                      display: 'flex',
                                      alignItems: 'center',
                                      justifyContent: 'center',
                                      width: '24px',
                                      height: '24px',
                                      userSelect: 'none'
                                    }}
                                  >
                                    +
                                  </button>
                                  <span style={{ fontSize: '12px', color: 'var(--text-secondary)', marginLeft: '4px', minWidth: '35px', textAlign: 'left' }}>
                                    {product.unit}
                                  </span>
                                </div>
                              </td>
                              <td style={{ padding: '12px', textAlign: 'right' }}>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleDeleteProduct(product.id);
                                  }}
                                  title={language === 'uz' ? "O'chirish" : "Удалить"}
                                  style={{
                                    border: 'none',
                                    backgroundColor: 'transparent',
                                    color: 'var(--warning-color)',
                                    padding: '4px',
                                    cursor: 'pointer',
                                    borderRadius: '4px',
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    justifyContent: 'center'
                                  }}
                                >
                                  <Trash2 size={16} />
                                </button>
                              </td>
                            </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              ) : (
                <>
                  {/* Agent Stock Table */}
                  <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                      <div>
                        <h3 style={{ fontSize: '16px', fontWeight: '500' }}>{language === 'uz' ? 'Agentlardagi jami mahsulot qoldig\'i' : 'Общий остаток товаров у агентов'}</h3>
                        <p style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>{language === 'uz' ? 'Kunlik biriktirilgan mahsulotlarning sotuvdan keyingi agentlardagi joriy qoldig\'i' : 'Текущий остаток привязанных товаров у агентов после продаж'}</p>
                      </div>
                      <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{t('all_products_count', { count: products.length })}</span>
                    </div>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                      <thead>
                        <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                          <th style={{ padding: '12px' }}>{t('barcode')}</th>
                          <th style={{ padding: '12px' }}>{t('product_name')}</th>
                          <th style={{ padding: '12px' }}>{t('selling_price')}</th>
                          <th style={{ padding: '12px', textAlign: 'center' }}>{t('assigned_agents')}</th>
                          <th style={{ padding: '12px', textAlign: 'right' }}>{t('total_assigned')}</th>
                          <th style={{ padding: '12px', textAlign: 'right' }}>{t('currently_remaining')}</th>
                          <th style={{ padding: '12px', textAlign: 'right' }}>{t('action')}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {products.map(product => {
                          const relevantAssignments = assignments.filter(ass => ass.productName === product.name);
                          const agentsCount = relevantAssignments.length;
                          const totalAssigned = relevantAssignments.reduce((sum, ass) => sum + ass.qty, 0);
                          const totalRemaining = relevantAssignments.reduce((sum, ass) => sum + (ass.remainingQty === "" ? 0 : (ass.remainingQty !== undefined ? ass.remainingQty : ass.qty)), 0);

                          return (
                            <tr 
                              key={product.id} 
                              onClick={() => {
                                if (agentsCount > 0) {
                                  setSelectedProductStockDetails(product.name);
                                }
                              }}
                              style={{ 
                                borderBottom: '1px solid var(--border-color)', 
                                fontSize: '14px',
                                cursor: agentsCount > 0 ? 'pointer' : 'default',
                                backgroundColor: 'transparent',
                                transition: 'background-color 0.2s'
                              }}
                              className={agentsCount > 0 ? "hover-row" : ""}
                            >
                              <td style={{ padding: '12px', fontFamily: 'monospace', color: 'var(--text-secondary)' }}>{product.barcode}</td>
                              <td style={{ padding: '12px', fontWeight: '500' }}>{product.name}</td>
                              <td style={{ padding: '12px', fontWeight: '600' }}>{product.price.toLocaleString()} UZS</td>
                              <td style={{ padding: '12px', textAlign: 'center' }}>
                                <span style={{ 
                                  backgroundColor: agentsCount > 0 ? 'rgba(99, 102, 241, 0.15)' : 'var(--bg-tertiary)', 
                                  color: agentsCount > 0 ? '#6366f1' : 'var(--text-secondary)', 
                                  padding: '3px 8px', 
                                  borderRadius: '4px', 
                                  fontSize: '12px',
                                  fontWeight: '600'
                                }}>
                                  {agentsCount} {language === 'uz' ? 'ta agent' : 'аг.'}
                                </span>
                              </td>
                              <td style={{ padding: '12px', textAlign: 'right', fontWeight: '700', color: 'var(--accent-color)' }}>
                                {totalAssigned} {product.unit}
                              </td>
                              <td style={{ padding: '12px', textAlign: 'right', fontWeight: '700', color: totalRemaining > 0 ? 'var(--success-color)' : 'var(--text-secondary)' }}>
                                {totalRemaining} {product.unit}
                              </td>
                              <td style={{ padding: '12px', textAlign: 'right' }}>
                                {agentsCount > 0 ? (
                                  <button 
                                    style={{ 
                                      padding: '4px 8px', 
                                      borderRadius: '6px', 
                                      border: '1px solid var(--accent-border)', 
                                      backgroundColor: 'var(--accent-light)', 
                                      color: 'var(--accent-color)', 
                                      fontSize: '11px', 
                                      fontWeight: '600',
                                      cursor: 'pointer' 
                                    }}
                                  >
                                    {language === 'uz' ? 'Ko\'rish' : 'Смотреть'}
                                  </button>
                                ) : (
                                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Mavjud emas' : 'Нет в наличии'}</span>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              )}

              {/* Add Product Modal Overlay */}
              {showAddProductModal && (
                <div style={{
                  position: 'fixed',
                  top: 0,
                  left: 0,
                  width: '100vw',
                  height: '100vh',
                  backgroundColor: 'rgba(15, 23, 42, 0.75)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 1000
                }} className="fade-in">
                  <div style={{
                    width: '460px',
                    backgroundColor: 'var(--bg-secondary)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: 'var(--shadow-lg)'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '24px' }}>{t('add_product_title')}</h3>
                    
                    <form onSubmit={handleAddProduct} style={{ display: 'flex', flexDirection: 'column', gap: '18px' }}>
                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('barcode')}</label>
                        <input 
                          type="text" 
                          required
                          value={newProduct.barcode}
                          onChange={(e) => setNewProduct({ ...newProduct, barcode: e.target.value })}
                          placeholder={language === 'uz' ? 'Shtrix kodni kiriting yoki skanerlang' : 'Введите или отсканируйте штрих-код'}
                          style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('product_name_label')}</label>
                        <input 
                          type="text" 
                          required
                          value={newProduct.name}
                          onChange={(e) => setNewProduct({ ...newProduct, name: e.target.value })}
                          placeholder={language === 'uz' ? 'Masalan: IQOS Iluma Pebble Grey' : 'Например: IQOS Iluma Pebble Grey'}
                          style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('price_label')}</label>
                          <input 
                            type="number" 
                            required
                            value={newProduct.price}
                            onChange={(e) => setNewProduct({ ...newProduct, price: e.target.value })}
                            placeholder={t('selling_price')}
                            style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('original_price_label')}</label>
                          <input 
                            type="number" 
                            value={newProduct.originalPrice}
                            onChange={(e) => setNewProduct({ ...newProduct, originalPrice: e.target.value })}
                            placeholder={t('original_price')}
                            style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('initial_stock_label')}</label>
                          <input 
                            type="number" 
                            required
                            value={newProduct.stock}
                            onChange={(e) => setNewProduct({ ...newProduct, stock: e.target.value })}
                            placeholder={language === 'uz' ? 'Omborda nechta bor?' : 'Сколько единиц на складе?'}
                            style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('unit_label')}</label>
                          <select 
                            value={newProduct.unit}
                            onChange={(e) => setNewProduct({ ...newProduct, unit: e.target.value })}
                            style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          >
                            <option value="dona">{t('dona')}</option>
                            <option value="blok">{t('blok')}</option>
                            <option value="quti">{t('quti')}</option>
                          </select>
                        </div>
                      </div>

                      <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                        <button 
                          type="button"
                          onClick={() => setShowAddProductModal(false)}
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            color: 'var(--text-secondary)',
                            fontWeight: '600',
                            fontSize: '14px',
                            cursor: 'pointer'
                          }}
                        >
                          {t('cancel')}
                        </button>
                        <button 
                          type="submit"
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: 'none',
                            backgroundColor: 'var(--accent-color)',
                            color: '#fff',
                            fontWeight: '600',
                            fontSize: '14px',
                            cursor: 'pointer'
                          }}
                        >
                          {t('save')}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* VIEW 3: STORES */}
          {activeTab === 'stores' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              
              {/* Stores Table Container (Full Width) */}
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
                  <h3 style={{ fontSize: '16px', fontWeight: '500' }}>{t('stores_list')}</h3>
                  <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                    {selectedStoreIds.length > 0 && (
                      <button
                        onClick={handleDeleteSelectedStores}
                        style={{
                          padding: '8px 16px',
                          borderRadius: '8px',
                          backgroundColor: 'var(--danger-color)',
                          color: '#fff',
                          fontSize: '13px',
                          fontWeight: '600',
                          border: 'none',
                          cursor: 'pointer',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '6px'
                        }}
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2v-6M10 11v6M14 11v6"/></svg>
                        {language === 'uz' ? `Tanlanganlarni o'chirish (${selectedStoreIds.length})` : `Удалить выбранные (${selectedStoreIds.length})`}
                      </button>
                    )}
                    
                    {/* Compact Excel Import */}
                    <label style={{
                      padding: '8px 10px',
                      borderRadius: '8px',
                      backgroundColor: 'rgba(13, 148, 136, 0.15)',
                      color: 'var(--accent-color)',
                      border: '1px solid rgba(13, 148, 136, 0.3)',
                      cursor: 'pointer',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      transition: 'all 0.2s',
                    }} title={t('import_excel')}>
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
                      <input 
                        type="file" 
                        accept=".xlsx" 
                        onChange={handleImportExcel} 
                        style={{ display: 'none' }} 
                      />
                    </label>

                    {/* Compact Excel Export */}
                    <button
                      onClick={handleExportExcel}
                      style={{
                        padding: '8px 10px',
                        borderRadius: '8px',
                        backgroundColor: 'rgba(13, 148, 136, 0.15)',
                        color: 'var(--accent-color)',
                        border: '1px solid rgba(13, 148, 136, 0.3)',
                        cursor: 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        transition: 'all 0.2s',
                      }}
                      title={language === 'uz' ? 'Excelga yuklash' : 'Экспорт в Excel'}
                    >
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>
                    </button>

                    {/* Add Store Button */}
                    <button
                      onClick={() => setShowAddStoreModal(true)}
                      style={{
                        padding: '8px 16px',
                        borderRadius: '8px',
                        backgroundColor: 'var(--accent-color)',
                        color: '#fff',
                        fontSize: '13px',
                        fontWeight: '600',
                        border: 'none',
                        cursor: 'pointer',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '6px',
                        transition: 'all 0.2s',
                      }}
                    >
                      <Plus size={16} />
                      {language === 'uz' ? "Do'kon qo'shish" : "Добавить магазин"}
                    </button>
                  </div>
                </div>

                  <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                    <thead>
                      <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                        <th style={{ padding: '12px', width: '40px' }}>
                          <input 
                            type="checkbox"
                            checked={stores.length > 0 && selectedStoreIds.length === stores.length}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setSelectedStoreIds(stores.map(s => s.id));
                              } else {
                                setSelectedStoreIds([]);
                              }
                            }}
                            style={{ cursor: 'pointer' }}
                          />
                        </th>
                        <th style={{ padding: '12px' }}>{language === 'uz' ? 'Do\'kon (Kodi & Nomi)' : 'Магазин (Код и Название)'}</th>
                        <th style={{ padding: '12px' }}>{language === 'uz' ? 'Manzili & Geopozitsiya' : 'Адрес и Геопозиция'}</th>
                        <th style={{ padding: '12px' }}>{language === 'uz' ? 'Yo\'nalish (SDE / Route)' : 'Маршрут (SDE / Route)'}</th>
                        <th style={{ padding: '12px' }}>{language === 'uz' ? 'Kategoriya / Status' : 'Категория / Статус'}</th>
                        <th style={{ padding: '12px', textAlign: 'center' }}>{t('actions')}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {stores.map((store) => {
                        const isChecked = selectedStoreIds.includes(store.id);
                        return (
                          <tr key={store.id} style={{ 
                            borderBottom: '1px solid var(--border-color)', 
                            fontSize: '14px',
                            backgroundColor: isChecked ? 'rgba(13, 148, 136, 0.05)' : 'transparent',
                            transition: 'background-color 0.2s'
                          }}>
                            <td style={{ padding: '12px' }}>
                              <input 
                                type="checkbox"
                                checked={isChecked}
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    setSelectedStoreIds([...selectedStoreIds, store.id]);
                                  } else {
                                    setSelectedStoreIds(selectedStoreIds.filter(id => id !== store.id));
                                  }
                                }}
                                style={{ cursor: 'pointer' }}
                              />
                            </td>
                            <td style={{ padding: '12px' }}>
                              <div style={{ fontWeight: '600', color: 'var(--text-primary)' }}>{store.name}</div>
                              <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>ID: {store.id}</span>
                            </td>
                            <td style={{ padding: '12px' }}>
                              <div style={{ fontSize: '13px', color: 'var(--text-primary)' }}>{store.address}</div>
                              {store.latitude && store.longitude ? (
                                <span style={{ fontSize: '11px', color: 'var(--accent-color)' }}>
                                  {store.latitude}, {store.longitude}
                                </span>
                              ) : null}
                            </td>
                            <td style={{ padding: '12px' }}>
                              <div style={{ fontWeight: '500' }}>{store.route}</div>
                              <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>{store.sde}</span>
                            </td>
                            <td style={{ padding: '12px' }}>
                              <div style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{store.census_type}</div>
                              <span style={{ 
                                fontSize: '11px', 
                                padding: '2px 6px', 
                                borderRadius: '4px',
                                backgroundColor: store.status ? 'rgba(239, 68, 110, 0.1)' : 'rgba(13, 148, 136, 0.1)',
                                color: store.status ? 'var(--danger-color)' : 'var(--accent-color)',
                                fontWeight: '600'
                              }}>
                                {store.status ? (language === 'uz' ? 'Xato RET' : 'Ошибка RET') : (language === 'uz' ? 'Normal' : 'Нормальный')}
                              </span>
                            </td>
                            <td style={{ padding: '12px', textAlign: 'center' }}>
                              <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                <a 
                                  href={store.map_link} 
                                  target="_blank" 
                                  rel="noreferrer"
                                  title={language === 'uz' ? 'Xaritada ochish' : 'Открыть на карте'}
                                  style={{
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    width: '28px',
                                    height: '28px',
                                    borderRadius: '6px',
                                    backgroundColor: 'var(--bg-primary)',
                                    border: '1px solid var(--border-color)',
                                    color: 'var(--accent-color)'
                                  }}
                                >
                                  <MapPin size={14} />
                                </a>
                                <button
                                  onClick={() => {
                                    setEditingStore({ ...store });
                                    setShowEditStoreModal(true);
                                  }}
                                  title={t('edit')}
                                  style={{
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    width: '28px',
                                    height: '28px',
                                    borderRadius: '6px',
                                    backgroundColor: 'var(--bg-primary)',
                                    border: '1px solid var(--border-color)',
                                    color: 'var(--text-primary)',
                                    cursor: 'pointer'
                                  }}
                                >
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>
                                </button>
                                <button
                                  onClick={() => handleDeleteStore(store.id)}
                                  title={t('delete')}
                                  style={{
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    width: '28px',
                                    height: '28px',
                                    borderRadius: '6px',
                                    backgroundColor: 'var(--bg-primary)',
                                    border: '1px solid var(--border-color)',
                                    color: 'var(--danger-color)',
                                    cursor: 'pointer'
                                  }}
                                >
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M10 11v6M14 11v6"/></svg>
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

              {/* Modal 4: Add Store */}
              {showAddStoreModal && (
                <div style={{
                  position: 'fixed',
                  top: 0,
                  left: 0,
                  width: '100vw',
                  height: '100vh',
                  backgroundColor: 'rgba(15, 23, 42, 0.75)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 1000
                }} className="fade-in">
                  <div style={{
                    width: '560px',
                    backgroundColor: 'var(--bg-secondary)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: 'var(--shadow-lg)',
                    maxHeight: '90vh',
                    overflowY: 'auto'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '20px' }}>{t('add_store_title')}</h3>
                    <form onSubmit={handleAddStore} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{language === 'uz' ? 'Do\'kon kodi (ID)' : 'Код магазина (ID)'}</label>
                          <input 
                            type="number" 
                            value={newStore.id}
                            onChange={(e) => setNewStore({ ...newStore, id: e.target.value })}
                            placeholder={language === 'uz' ? 'Avtomatik (ixtiyoriy)' : 'Автоматически (опционально)'}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{language === 'uz' ? 'Do\'kon nomi' : 'Название магазина'}</label>
                          <input 
                            type="text" 
                            required
                            value={newStore.name}
                            onChange={(e) => setNewStore({ ...newStore, name: e.target.value })}
                            placeholder={language === 'uz' ? 'Masalan: Humo smoke' : 'Например: Humo smoke'}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('owner_name_label')}</label>
                          <input 
                            type="text" 
                            value={newStore.owner_name}
                            onChange={(e) => setNewStore({ ...newStore, owner_name: e.target.value })}
                            placeholder={language === 'uz' ? 'F.I.SH.' : 'Ф.И.О.'}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('phone_label')}</label>
                          <input 
                            type="text" 
                            value={newStore.phone}
                            onChange={(e) => setNewStore({ ...newStore, phone: e.target.value })}
                            placeholder="+998 90..."
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('address_label')}</label>
                        <input 
                          type="text" 
                          required
                          value={newStore.address}
                          onChange={(e) => setNewStore({ ...newStore, address: e.target.value })}
                          placeholder={language === 'uz' ? 'Chilonzor 4-daha...' : 'Чиланзар 4-квартал...'}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('map_link_label')}</label>
                        <input 
                          type="text" 
                          value={newStore.map_link}
                          onChange={(e) => handleMapLinkChange(e.target.value, false)}
                          placeholder={language === 'uz' ? 'Havolani kiritsangiz koordinatalar avtomatik to\'ladi' : 'Вставьте ссылку, чтобы координаты заполнились автоматически'}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px dashed var(--accent-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('lat_label')}</label>
                          <input 
                            type="text" 
                            value={newStore.latitude}
                            onChange={(e) => setNewStore({ ...newStore, latitude: e.target.value })}
                            placeholder="41.0015"
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('lng_label')}</label>
                          <input 
                            type="text" 
                            value={newStore.longitude}
                            onChange={(e) => setNewStore({ ...newStore, longitude: e.target.value })}
                            placeholder="71.5865"
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('route_label')}</label>
                          <input 
                            type="text" 
                            value={newStore.route}
                            onChange={(e) => setNewStore({ ...newStore, route: e.target.value, sde: `SDE_${e.target.value}` })}
                            placeholder="NM_01"
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('status_label')}</label>
                          <select 
                            value={newStore.status}
                            onChange={(e) => setNewStore({ ...newStore, status: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          >
                            <option value="0">{language === 'uz' ? '0 (Normal)' : '0 (Нормальный)'}</option>
                            <option value="1">{language === 'uz' ? '1 (Xato RET)' : '1 (Ошибка RET)'}</option>
                          </select>
                        </div>
                      </div>

                      <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                        <button 
                          type="button"
                          onClick={() => setShowAddStoreModal(false)}
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            color: 'var(--text-secondary)',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('cancel')}
                        </button>
                        <button 
                          type="submit"
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: 'none',
                            backgroundColor: 'var(--accent-color)',
                            color: '#fff',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('save')}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}

              {/* Modal 3: Edit Store */}
              {showEditStoreModal && editingStore && (
                <div style={{
                  position: 'fixed',
                  top: 0,
                  left: 0,
                  width: '100vw',
                  height: '100vh',
                  backgroundColor: 'rgba(15, 23, 42, 0.75)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 1000
                }} className="fade-in">
                  <div style={{
                    width: '560px',
                    backgroundColor: 'var(--bg-secondary)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: 'var(--shadow-lg)'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '20px' }}>{language === 'uz' ? 'Do\'konni tahrirlash' : 'Редактировать магазин'}</h3>
                    <form onSubmit={handleUpdateStore} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{language === 'uz' ? 'Do\'kon kodi (ID)' : 'Код магазина (ID)'}</label>
                          <input 
                            type="number" 
                            disabled
                            value={editingStore.id}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-secondary)', fontSize: '14px', cursor: 'not-allowed' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{language === 'uz' ? 'Do\'kon nomi' : 'Название магазина'}</label>
                          <input 
                            type="text" 
                            required
                            value={editingStore.name}
                            onChange={(e) => setEditingStore({ ...editingStore, name: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('owner_name_label')}</label>
                          <input 
                            type="text" 
                            value={editingStore.owner_name}
                            onChange={(e) => setEditingStore({ ...editingStore, owner_name: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('phone_label')}</label>
                          <input 
                            type="text" 
                            value={editingStore.phone}
                            onChange={(e) => setEditingStore({ ...editingStore, phone: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('address_label')}</label>
                        <input 
                          type="text" 
                          required
                          value={editingStore.address}
                          onChange={(e) => setEditingStore({ ...editingStore, address: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px', fontWeight: '500' }}>{t('map_link_label')}</label>
                        <input 
                          type="text" 
                          value={editingStore.map_link || ''}
                          onChange={(e) => handleMapLinkChange(e.target.value, true)}
                          placeholder={language === 'uz' ? 'Havolani kiritsangiz koordinatalar avtomatik yangilanadi' : 'Вставьте ссылку, чтобы координаты обновились автоматически'}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px dashed var(--accent-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('lat_label')}</label>
                          <input 
                            type="text" 
                            value={editingStore.latitude}
                            onChange={(e) => setEditingStore({ ...editingStore, latitude: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('lng_label')}</label>
                          <input 
                            type="text" 
                            value={editingStore.longitude}
                            onChange={(e) => setEditingStore({ ...editingStore, longitude: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('route_label')}</label>
                          <input 
                            type="text" 
                            value={editingStore.route}
                            onChange={(e) => setEditingStore({ ...editingStore, route: e.target.value, sde: `SDE_${e.target.value}` })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          />
                        </div>
                        <div>
                          <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('status_label')}</label>
                          <select 
                            value={editingStore.status}
                            onChange={(e) => setEditingStore({ ...editingStore, status: e.target.value })}
                            style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                          >
                            <option value="0">{language === 'uz' ? '0 (Normal)' : '0 (Нормальный)'}</option>
                            <option value="1">{language === 'uz' ? '1 (Xato RET)' : '1 (Ошибка RET)'}</option>
                          </select>
                        </div>
                      </div>

                      <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                        <button 
                          type="button"
                          onClick={() => {
                            setEditingStore(null);
                            setShowEditStoreModal(false);
                          }}
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            color: 'var(--text-secondary)',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('cancel')}
                        </button>
                        <button 
                          type="submit"
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: 'none',
                            backgroundColor: 'var(--accent-color)',
                            color: '#fff',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('save')}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}

            </div>
          )}

          {/* VIEW: AGENT MANAGEMENT */}
          {(activeTab === 'agents' || activeTab === 'settings_agents') && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 2fr', gap: '24px' }}>
                
                {/* Add Agent Form */}
                <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', height: 'fit-content' }}>
                  <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '20px' }}>{t('add_agent_title')}</h3>
                  <form onSubmit={handleAddAgent} style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {t('agent_login_label')}
                      </label>
                      <input 
                        type="text" 
                        required
                        value={newAgent.login}
                        onChange={(e) => setNewAgent({ ...newAgent, login: e.target.value })}
                        placeholder={language === 'uz' ? 'Masalan: AGENT-FG-R3' : 'Например: AGENT-FG-R3'}
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                      />
                    </div>

                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {t('agent_password_label')}
                      </label>
                      <input 
                        type="text" 
                        required
                        value={newAgent.password}
                        onChange={(e) => setNewAgent({ ...newAgent, password: e.target.value })}
                        placeholder={language === 'uz' ? 'Parol kiriting' : 'Введите пароль'}
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>

                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {t('agent_username_label')}
                      </label>
                      <input 
                        type="text" 
                        value={newAgent.username}
                        onChange={(e) => setNewAgent({ ...newAgent, username: e.target.value })}
                        placeholder="sherzod_agent"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>

                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {t('agent_name_label')}
                      </label>
                      <input 
                        type="text" 
                        value={newAgent.name}
                        onChange={(e) => setNewAgent({ ...newAgent, name: e.target.value })}
                        placeholder={language === 'uz' ? 'Sherzod Alimov' : 'Шерзод Алимов'}
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>

                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {t('agent_phone_label')}
                      </label>
                      <input 
                        type="text" 
                        value={newAgent.phone}
                        onChange={(e) => setNewAgent({ ...newAgent, phone: e.target.value })}
                        placeholder="+998 90..."
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>

                    <button 
                      type="submit"
                      style={{
                        padding: '11px',
                        borderRadius: '8px',
                        border: 'none',
                        backgroundColor: 'var(--accent-color)',
                        color: '#fff',
                        fontSize: '14px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: '8px',
                        marginTop: '8px'
                      }}
                    >
                      <Plus size={16} /> {t('add_agent')}
                    </button>
                  </form>
                </div>

                {/* Agents List Table */}
                <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                  <h3 style={{ fontSize: '16px', fontWeight: '500' }}>{t('agents_list')}</h3>
                  
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                      <thead>
                        <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                          <th style={{ padding: '12px' }}>ID</th>
                          <th style={{ padding: '12px' }}>{language === 'uz' ? 'Login / Parol' : 'Логин / Пароль'}</th>
                          <th style={{ padding: '12px' }}>{language === 'uz' ? 'Ismi & Username' : 'Имя и Username'}</th>
                          <th style={{ padding: '12px' }}>{t('phone')}</th>
                          <th style={{ padding: '12px', textAlign: 'center' }}>{t('actions')}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {agents.map((agent) => (
                          <tr key={agent.id} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}>
                            <td style={{ padding: '12px', color: 'var(--text-muted)' }}>{agent.id}</td>
                            <td style={{ padding: '12px' }}>
                              <div style={{ fontWeight: '700', color: 'var(--text-primary)' }}>{agent.login || `AGENT-FG-R${agent.id}`}</div>
                              <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>{t('password')}: <span style={{ fontFamily: 'monospace', fontWeight: '600', backgroundColor: 'var(--bg-tertiary)', padding: '1px 4px', borderRadius: '4px' }}>{agent.password || '123'}</span></span>
                            </td>
                            <td style={{ padding: '12px' }}>
                              <div style={{ fontWeight: '500' }}>{agent.name || '-'}</div>
                              <span style={{ fontSize: '11px', color: 'var(--accent-color)' }}>@{agent.username || t('no_username')}</span>
                            </td>
                            <td style={{ padding: '12px', color: 'var(--text-primary)' }}>
                              {agent.phone || '-'}
                            </td>
                            <td style={{ padding: '12px', textAlign: 'center' }}>
                              <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', alignItems: 'center' }}>
                                <button
                                  onClick={() => {
                                    setEditingAgent({ ...agent });
                                    setShowEditAgentModal(true);
                                  }}
                                  title={t('edit')}
                                  style={{
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    width: '28px',
                                    height: '28px',
                                    borderRadius: '6px',
                                    backgroundColor: 'var(--bg-primary)',
                                    border: '1px solid var(--border-color)',
                                    color: 'var(--text-primary)',
                                    cursor: 'pointer'
                                  }}
                                >
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>
                                </button>
                                <button
                                  onClick={() => handleDeleteAgent(agent.id)}
                                  title={t('delete')}
                                  style={{
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    width: '28px',
                                    height: '28px',
                                    borderRadius: '6px',
                                    backgroundColor: 'var(--bg-primary)',
                                    border: '1px solid var(--border-color)',
                                    color: 'var(--danger-color)',
                                    cursor: 'pointer'
                                  }}
                                >
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M10 11v6M14 11v6"/></svg>
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

              </div>

              {/* Edit Agent Modal */}
              {showEditAgentModal && editingAgent && (
                <div style={{
                  position: 'fixed',
                  top: 0,
                  left: 0,
                  width: '100vw',
                  height: '100vh',
                  backgroundColor: 'rgba(15, 23, 42, 0.75)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 1000
                }} className="fade-in">
                  <div style={{
                    width: '480px',
                    backgroundColor: 'var(--bg-secondary)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: 'var(--shadow-lg)'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '20px' }}>{language === 'uz' ? 'Agent ma\'lumotlarini tahrirlash' : 'Редактировать данные агента'}</h3>
                    <form onSubmit={handleUpdateAgent} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('agent_login_label')}</label>
                        <input 
                          type="text" 
                          required
                          value={editingAgent.login || ''}
                          onChange={(e) => setEditingAgent({ ...editingAgent, login: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', fontWeight: '600' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('agent_password_label')}</label>
                        <input 
                          type="text" 
                          required
                          value={editingAgent.password || ''}
                          onChange={(e) => setEditingAgent({ ...editingAgent, password: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('agent_username_label')}</label>
                        <input 
                          type="text" 
                          value={editingAgent.username || ''}
                          onChange={(e) => setEditingAgent({ ...editingAgent, username: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('agent_name_label')}</label>
                        <input 
                          type="text" 
                          value={editingAgent.name || ''}
                          onChange={(e) => setEditingAgent({ ...editingAgent, name: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('agent_phone_label')}</label>
                        <input 
                          type="text" 
                          value={editingAgent.phone || ''}
                          onChange={(e) => setEditingAgent({ ...editingAgent, phone: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                        <button 
                          type="button"
                          onClick={() => {
                            setEditingAgent(null);
                            setShowEditAgentModal(false);
                          }}
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            color: 'var(--text-secondary)',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('cancel')}
                        </button>
                        <button 
                          type="submit"
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: 'none',
                            backgroundColor: 'var(--accent-color)',
                            color: '#fff',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('save')}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}

              {/* Edit Admin Modal */}
              {showEditAdminModal && editingAdmin && (
                <div style={{
                  position: 'fixed',
                  top: 0,
                  left: 0,
                  width: '100vw',
                  height: '100vh',
                  backgroundColor: 'rgba(15, 23, 42, 0.75)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 1000
                }} className="fade-in">
                  <div style={{
                    width: '480px',
                    backgroundColor: 'var(--bg-secondary)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: 'var(--shadow-lg)'
                  }}>
                    <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '20px', color: 'var(--text-primary)' }}>
                      {language === 'uz' ? 'Administrator ma\'lumotlarini tahrirlash' : 'Редактировать данные администратора'}
                    </h3>
                    <form onSubmit={handleUpdateAdmin} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                          {language === 'uz' ? "Login (Username)" : "Логин (Username)"}
                        </label>
                        <input 
                          type="text" 
                          required
                          value={editingAdmin.login || ''}
                          onChange={(e) => setEditingAdmin({ ...editingAdmin, login: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', fontWeight: '600' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                          {language === 'uz' ? "Yangi Parol (agar o'zgartirmoqchi bo'lsangiz)" : "Новый Пароль (если хотите изменить)"}
                        </label>
                        <input 
                          type="password" 
                          value={editingAdmin.password || ''}
                          onChange={(e) => setEditingAdmin({ ...editingAdmin, password: e.target.value })}
                          placeholder={language === 'uz' ? "O'zgarishsiz qoldirish uchun bo'sh qoldiring" : "Оставьте пустым для сохранения прежнего"}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                          {language === 'uz' ? "Ism (F.I.SH.)" : "Имя (Ф.И.О.)"}
                        </label>
                        <input 
                          type="text" 
                          value={editingAdmin.name || ''}
                          onChange={(e) => setEditingAdmin({ ...editingAdmin, name: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                          {language === 'uz' ? "Telefon" : "Телефон"}
                        </label>
                        <input 
                          type="text" 
                          value={editingAdmin.phone || ''}
                          onChange={(e) => setEditingAdmin({ ...editingAdmin, phone: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                        />
                      </div>

                      <div>
                        <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                          {language === 'uz' ? "Rol" : "Роль"}
                        </label>
                        <select 
                          value={editingAdmin.role || 'admin'}
                          onChange={(e) => setEditingAdmin({ ...editingAdmin, role: e.target.value })}
                          style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', cursor: 'pointer' }}
                        >
                          <option value="admin">{language === 'uz' ? "Bosh Admin (Administrator)" : "Главный Админ (Администратор)"}</option>
                          <option value="warehouse_manager">{language === 'uz' ? "Ombor Mudiri" : "Заведующий складом"}</option>
                        </select>
                      </div>

                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px', margin: '6px 0' }}>
                        <input 
                          type="checkbox"
                          id="edit-admin-active"
                          checked={editingAdmin.is_active !== false}
                          onChange={(e) => setEditingAdmin({ ...editingAdmin, is_active: e.target.checked })}
                          style={{ width: '16px', height: '16px', cursor: 'pointer' }}
                        />
                        <label htmlFor="edit-admin-active" style={{ fontSize: '13px', color: 'var(--text-primary)', cursor: 'pointer', fontWeight: '500' }}>
                          {language === 'uz' ? "Faol (Tizimga kira oladi)" : "Активен (Может входить в систему)"}
                        </label>
                      </div>

                      <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                        <button 
                          type="button"
                          onClick={() => {
                            setEditingAdmin(null);
                            setShowEditAdminModal(false);
                          }}
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            color: 'var(--text-secondary)',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('cancel')}
                        </button>
                        <button 
                          type="submit"
                          style={{
                            flex: 1,
                            padding: '12px',
                            borderRadius: '8px',
                            border: 'none',
                            backgroundColor: 'var(--accent-color)',
                            color: '#fff',
                            fontWeight: '600',
                            cursor: 'pointer'
                          }}
                        >
                          {t('save')}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}

            </div>
          )}

          {/* VIEW 4: AGENT ASSIGNMENT */}
          {activeTab === 'assignments' && (
            <div className="fade-in" style={{ display: 'flex', gap: '24px', height: '100%', alignItems: 'stretch' }}>
              
              {/* Left Column: Agent List */}
              {userRole !== 'agent' && (
                <div style={{
                  width: '300px',
                  backgroundColor: 'var(--bg-secondary)',
                  border: '1px solid var(--border-color)',
                  borderRadius: '12px',
                  padding: '20px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '16px'
                }}>
                  <h3 style={{ fontSize: '16px', fontWeight: '600' }}>{language === 'uz' ? 'Agentlar ro\'yxati' : 'Список агентов'}</h3>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', overflowY: 'auto', flexGrow: 1 }}>
                    {agents.map(agent => (
                      <div 
                      key={agent.id}
                      onClick={() => setSelectedAgentId(agent.id)}
                      style={{
                        padding: '12px 16px',
                        borderRadius: '8px',
                        border: '1px solid',
                        borderColor: selectedAgentId === agent.id ? 'var(--accent-color)' : 'var(--border-color)',
                        backgroundColor: selectedAgentId === agent.id ? 'var(--accent-light)' : 'var(--bg-primary)',
                        cursor: 'pointer',
                        transition: 'all var(--transition-fast)'
                      }}
                    >
                      <div style={{ fontSize: '14px', fontWeight: '600', color: selectedAgentId === agent.id ? 'var(--accent-color)' : 'var(--text-primary)' }}>
                        {agent.login}
                      </div>
                      <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>
                        {agent.username ? `@${agent.username.replace(/^@/, '')}` : (language === 'uz' ? 'Username kiritilmagan' : 'Username не указан')}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
              )}

              {/* Right Column: Selected Agent Assignment Details */}
              {(() => {
                const selectedAgent = userRole === 'agent'
                  ? { id: parseInt(currentUserId || localStorage.getItem('currentUserId') || '0'), login: adminName, username: adminName }
                  : agents.find(a => a.id === selectedAgentId);
                if (!selectedAgent) {
                  return (
                    <div style={{ flexGrow: 1, backgroundColor: 'var(--bg-secondary)', border: '1px solid var(--border-color)', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--text-secondary)' }}>
                      {language === 'uz' ? 'Iltimos, agentni tanlang' : 'Пожалуйста, выберите агента'}
                    </div>
                  );
                }

                // Filter assignments for this agent by agentId (highly secure and works for empty names)
                // (Now using top-level reactive variables: agentProducts and agentStores)

                return (
                  <div style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', gap: '24px' }}>
                    
                    {/* Header Details Card */}
                    <div style={{
                      backgroundColor: 'var(--bg-secondary)',
                      border: '1px solid var(--border-color)',
                      borderRadius: '12px',
                      padding: '24px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      flexWrap: 'wrap',
                      gap: '16px'
                    }}>
                      <div>
                        <div style={{ fontSize: '12px', color: 'var(--accent-color)', fontWeight: '600', textTransform: 'uppercase', marginBottom: '4px' }}>
                          {userRole === 'agent' ? (language === 'uz' ? "Sizga Biriktirilgan Ma'lumotlar" : "Закрепленные за вами данные") : (language === 'uz' ? "Faol Agent" : "Активный агент")}
                        </div>
                        <h2 style={{ fontSize: '20px', fontWeight: '700', margin: 0 }}>{selectedAgent.login}</h2>
                        <div style={{ fontSize: '13px', color: 'var(--text-secondary)', marginTop: '4px', display: 'flex', gap: '16px', flexWrap: 'wrap' }}>
                          {selectedAgent.login && <span>{language === 'uz' ? 'Login:' : 'Логин:'} {selectedAgent.login}</span>}
                          <span>Username: {selectedAgent.username ? `@${selectedAgent.username.replace(/^@/, '')}` : (language === 'uz' ? 'Kiritilmagan' : 'Не указано')}</span>
                          {selectedAgent.phone && <span>Tel: {selectedAgent.phone}</span>}
                        </div>
                      </div>

                      {userRole !== 'agent' && (
                        <div style={{ display: 'flex', gap: '12px' }}>
                          <button 
                            onClick={() => {
                              setNewAssignment({ productId: '', qty: '' });
                              setShowAssignProductModal(true);
                            }}
                            style={{
                              padding: '10px 18px',
                              borderRadius: '8px',
                              border: 'none',
                              backgroundColor: 'var(--accent-color)',
                              color: '#fff',
                              fontSize: '13px',
                              fontWeight: '600',
                              cursor: 'pointer',
                              display: 'flex',
                              alignItems: 'center',
                              gap: '6px'
                            }}
                          >
                            <Plus size={14} /> {language === 'uz' ? 'Tovar biriktirish' : 'Закрепить товар'}
                          </button>
                          <button 
                            onClick={() => {
                              setNewStoreAssignment({ storeId: '' });
                              setShowAssignStoreModal(true);
                            }}
                            style={{
                              padding: '10px 18px',
                              borderRadius: '8px',
                              border: 'none',
                              backgroundColor: 'var(--accent-color)',
                              color: '#fff',
                              fontSize: '13px',
                              fontWeight: '600',
                              cursor: 'pointer',
                              display: 'flex',
                              alignItems: 'center',
                              gap: '6px'
                            }}
                          >
                            <Plus size={14} /> {language === 'uz' ? 'Do\'kon biriktirish' : 'Закрепить магазин'}
                          </button>
                        </div>
                      )}
                    </div>

                    {/* Lists Grid */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '24px', flexGrow: 1 }}>
                      
                      {/* Stores List Box */}
                      <div style={{
                        backgroundColor: 'var(--bg-secondary)',
                        border: '1px solid var(--border-color)',
                        borderRadius: '12px',
                        padding: '20px',
                        display: window.innerWidth > 768 ? 'flex' : 'block',
                        flexDirection: 'column',
                        width: '100%',
                        maxWidth: '100%',
                        overflowX: window.innerWidth > 768 ? 'visible' : 'auto',
                        minWidth: 0
                      }}>
                        <h3 style={{ fontSize: '15px', fontWeight: '600', marginBottom: '16px', color: 'var(--text-primary)' }}>{language === 'uz' ? "Biriktirilgan do'konlar (Kanal)" : 'Закрепленные магазины (Канал)'}</h3>
                        <div style={{ 
                          overflowX: userRole === 'agent' ? 'hidden' : 'auto', 
                          overflowY: 'auto', 
                          flexGrow: 1, 
                          width: '100%', 
                          maxWidth: '100%', 
                          minWidth: 0, 
                          display: 'block', 
                          WebkitOverflowScrolling: 'touch' 
                        }}>
                          {userRole === 'agent' ? (
                            activeAgentStores.length === 0 ? (
                              <div style={{ padding: '32px', textAlign: 'center', color: '#10b981', fontSize: '14px', fontWeight: '600' }}>
                                🎉 {language === 'uz' ? "Bugungi barcha vazifalar muvaffaqiyatli bajarildi!" : "Все сегодняшние задачи успешно выполнены!"}
                              </div>
                            ) : (
                              /* Agent View: Beautiful div-based list */
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '0', width: '100%' }}>
                                {activeAgentStores.map((store, idx) => (
                                  <div 
                                    key={store.id} 
                                    onClick={() => handleOpenCashier(store)}
                                    style={{ 
                                      display: 'flex', 
                                      alignItems: 'center', 
                                      gap: '10px', 
                                      padding: '14px 0', 
                                      borderBottom: '1px solid var(--border-color)',
                                      width: '100%',
                                      boxSizing: 'border-box',
                                      cursor: 'pointer'
                                    }}
                                    className="hoverable-row-div"
                                  >
                                    {/* Subtle index number */}
                                    <span style={{ 
                                      fontSize: '11px', 
                                      color: 'var(--text-muted)', 
                                      fontWeight: '600',
                                      minWidth: '16px',
                                      textAlign: 'center',
                                      flexShrink: 0
                                    }}>
                                      {idx + 1}.
                                    </span>

                                    {/* Map Link / Map Pin button */}
                                    {store.map_link || (store.latitude && store.longitude) ? (
                                      <a
                                        href={store.map_link || `https://maps.google.com/?q=${store.latitude},${store.longitude}`}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        style={{
                                          display: 'inline-flex',
                                          alignItems: 'center',
                                          justifyContent: 'center',
                                          width: '26px',
                                          height: '26px',
                                          borderRadius: '50%',
                                          backgroundColor: 'var(--accent-light)',
                                          color: 'var(--accent-color)',
                                          transition: 'all 0.2s ease',
                                          cursor: 'pointer',
                                          border: 'none',
                                          flexShrink: 0
                                        }}
                                        className="map-pin-btn"
                                        title={language === 'uz' ? "Xaritada ko'rish" : "Посмотреть на карте"}
                                        onClick={(e) => e.stopPropagation()}
                                      >
                                        <MapPin size={13} />
                                      </a>
                                    ) : (
                                      <span style={{ color: 'var(--text-muted)', fontSize: '11px', minWidth: '26px', textAlign: 'center', flexShrink: 0 }}>—</span>
                                    )}

                                    {/* Store name - small font, wraps to 2 lines max */}
                                    <span style={{ 
                                      fontSize: '12px', 
                                      fontWeight: '600', 
                                      color: 'var(--text-primary)',
                                      lineHeight: '1.3',
                                      wordBreak: 'break-word',
                                      display: '-webkit-box',
                                      WebkitLineClamp: 2,
                                      WebkitBoxOrient: 'vertical',
                                      overflow: 'hidden',
                                      textOverflow: 'ellipsis',
                                      flexGrow: 1
                                    }}>
                                      {store.storeName}
                                    </span>

                                    {/* Chevron icon indicator */}
                                    <span style={{ fontSize: '12px', color: 'var(--accent-color)', fontWeight: '600', flexShrink: 0, paddingRight: '8px' }}>
                                      {language === 'uz' ? "Kassa ➔" : "Касса ➔"}
                                    </span>
                                  </div>
                                ))}
                              </div>
                            )
                          ) : (
                            /* Admin View: Original drag-and-drop table layout */
                            agentStores.length === 0 ? (
                              <div style={{ padding: '24px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13px' }}>{language === 'uz' ? "Do'konlar biriktirilmagan" : 'Магазины не закреплены'}</div>
                            ) : (
                              <table style={{ width: '100%', minWidth: '600px', borderCollapse: 'collapse', textAlign: 'left' }}>
                                <thead>
                                  <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '12px' }}>
                                    <th style={{ padding: '8px 4px', width: '36px' }}>№</th>
                                    <th style={{ padding: '8px 4px' }}>{language === 'uz' ? "Do'kon / Rahbar / Manzil" : 'Магазин / Руководитель / Адрес'}</th>
                                    <th style={{ padding: '8px 4px', textAlign: 'center', width: '70px' }}>{language === 'uz' ? 'Lokatsiya' : 'Локация'}</th>
                                    <th style={{ padding: '8px 4px', textAlign: 'right', width: '60px' }}>{language === 'uz' ? "O'chirish" : 'Удалить'}</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {agentStores.map((store, idx) => (
                                    <tr 
                                      key={store.id} 
                                      draggable={window.innerWidth > 768}
                                      onDragStart={(e) => handleStoreDragStart(e, idx)}
                                      onDragOver={handleStoreDragOver}
                                      onDrop={(e) => handleStoreDrop(e, idx)}
                                      style={{ 
                                        borderBottom: '1px solid var(--border-color)', 
                                        fontSize: '13px', 
                                        cursor: window.innerWidth > 768 ? 'grab' : 'default',
                                        backgroundColor: draggedStoreIndex === idx ? 'var(--bg-primary)' : 'transparent',
                                        opacity: draggedStoreIndex === idx ? 0.5 : 1
                                      }}
                                      className="hoverable-row"
                                    >
                                      <td style={{ padding: '10px 4px' }}>
                                        <div style={{
                                          width: '22px',
                                          height: '22px',
                                          borderRadius: '50%',
                                          backgroundColor: 'var(--accent-color)',
                                          color: '#fff',
                                          display: 'flex',
                                          alignItems: 'center',
                                          justifyContent: 'center',
                                          fontSize: '11px',
                                          fontWeight: '700'
                                        }}>
                                          {idx + 1}
                                        </div>
                                      </td>
                                      <td style={{ padding: '10px 4px' }}>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                          {/* Drag handle for desktop */}
                                          <span style={{ color: 'var(--text-muted)', fontSize: '14px', cursor: 'grab', userSelect: 'none' }}>☰</span>
                                          
                                          {/* Up/Down buttons for mobile & easy desktop sorting */}
                                          <div style={{ display: 'flex', flexDirection: 'column', gap: '2px', alignItems: 'center' }}>
                                            <button
                                              onClick={(e) => {
                                                e.stopPropagation();
                                                handleMoveStore(idx, -1);
                                              }}
                                              disabled={idx === 0}
                                              style={{
                                                border: 'none',
                                                backgroundColor: 'transparent',
                                                color: idx === 0 ? 'var(--text-muted)' : 'var(--accent-color)',
                                                padding: '1px 3px',
                                                cursor: idx === 0 ? 'not-allowed' : 'pointer',
                                                display: 'inline-flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                fontSize: '9px',
                                                lineHeight: '1',
                                                fontWeight: 'bold'
                                              }}
                                              title={language === 'uz' ? "Tepaga ko'tarish" : "Поднять вверх"}
                                            >
                                              ▲
                                            </button>
                                            <button
                                              onClick={(e) => {
                                                e.stopPropagation();
                                                handleMoveStore(idx, 1);
                                              }}
                                              disabled={idx === agentStores.length - 1}
                                              style={{
                                                border: 'none',
                                                backgroundColor: 'transparent',
                                                color: idx === agentStores.length - 1 ? 'var(--text-muted)' : 'var(--accent-color)',
                                                padding: '1px 3px',
                                                cursor: idx === agentStores.length - 1 ? 'not-allowed' : 'pointer',
                                                display: 'inline-flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                fontSize: '9px',
                                                lineHeight: '1',
                                                fontWeight: 'bold'
                                              }}
                                              title={language === 'uz' ? "Pastga tushirish" : "Опустить вниз"}
                                            >
                                              ▼
                                            </button>
                                          </div>

                                          <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                                            <span style={{ fontWeight: '600', color: 'var(--text-primary)', fontSize: '14px' }}>{store.storeName}</span>
                                            <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                                              👤 {store.ownerName || 'Tadbirkor'} • 📞 {store.phone || "Telefon yo'q"}
                                            </span>
                                            {store.address && (
                                              <span style={{ fontSize: '11px', color: 'var(--text-secondary)', fontStyle: 'italic' }}>
                                                📍 {store.address}
                                              </span>
                                            )}
                                          </div>
                                        </div>
                                      </td>
                                      <td style={{ padding: '10px 4px', textAlign: 'center' }}>
                                        {store.map_link || (store.latitude && store.longitude) ? (
                                          <a
                                            href={store.map_link || `https://maps.google.com/?q=${store.latitude},${store.longitude}`}
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            style={{
                                              display: 'inline-flex',
                                              alignItems: 'center',
                                              justifyContent: 'center',
                                              width: '32px',
                                              height: '32px',
                                              borderRadius: '50%',
                                              backgroundColor: 'var(--accent-light)',
                                              color: 'var(--accent-color)',
                                              transition: 'all 0.2s ease',
                                              cursor: 'pointer',
                                              border: 'none'
                                            }}
                                            className="map-pin-btn"
                                            title={language === 'uz' ? "Xaritada ko'rish" : "Посмотреть на карте"}
                                            onClick={(e) => e.stopPropagation()}
                                          >
                                            <MapPin size={16} />
                                          </a>
                                        ) : (
                                          <span style={{ color: 'var(--text-muted)', fontSize: '11px' }}>—</span>
                                        )}
                                      </td>
                                      <td style={{ padding: '10px 4px', textAlign: 'right' }}>
                                        <button
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            handleDeleteStoreAssignment(store.id);
                                          }}
                                          title={language === 'uz' ? "O'chirish" : "Удалить"}
                                          style={{
                                            border: 'none',
                                            backgroundColor: 'transparent',
                                            color: 'var(--warning-color)',
                                            padding: '4px',
                                            cursor: 'pointer',
                                            borderRadius: '4px',
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            justifyContent: 'center'
                                          }}
                                        >
                                          <Trash2 size={15} />
                                        </button>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            )
                          )}
                        </div>
                      </div>

                      {/* Products List Box */}
                      <div style={{
                        backgroundColor: 'var(--bg-secondary)',
                        border: '1px solid var(--border-color)',
                        borderRadius: '12px',
                        padding: '20px',
                        display: window.innerWidth > 768 ? 'flex' : 'block',
                        flexDirection: 'column',
                        width: '100%',
                        maxWidth: '100%',
                        overflowX: window.innerWidth > 768 ? 'visible' : 'auto',
                        minWidth: 0
                      }}>
                        <h3 style={{ fontSize: '15px', fontWeight: '600', marginBottom: '16px', color: 'var(--text-primary)' }}>{language === 'uz' ? 'Biriktirilgan mahsulotlar (Bugun uchun)' : 'Закрепленные товары (на сегодня)'}</h3>
                        <div style={{ 
                          overflowX: userRole === 'agent' ? 'hidden' : 'auto', 
                          overflowY: 'auto', 
                          flexGrow: 1, 
                          width: '100%', 
                          maxWidth: '100%', 
                          minWidth: 0, 
                          display: 'block', 
                          WebkitOverflowScrolling: 'touch' 
                        }}>
                          {agentProducts.length === 0 ? (
                            <div style={{ padding: '24px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13px' }}>{language === 'uz' ? 'Mahsulotlar biriktirilmagan' : 'Товары не закреплены'}</div>
                          ) : userRole === 'agent' ? (
                            /* Agent View: Beautiful div-based list (no tables to prevent horizontal overflow/scroll) */
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '0', width: '100%' }}>
                              <div style={{ display: 'flex', gap: '8px', paddingBottom: '8px', borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '11px', fontWeight: '600', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                                <span>{language === 'uz' ? 'Mahsulot va Qoldiq' : 'Товар и Остаток'}</span>
                              </div>
                              {agentProducts.map(prod => (
                                <div 
                                  key={prod.id}
                                  onClick={() => {
                                    setSelectedAssignmentForEdit(prod);
                                    setEditAssignmentQty(prod.qty.toString());
                                    setEditAssignmentRemainingQty((prod.remainingQty !== undefined ? prod.remainingQty : prod.qty).toString());
                                    setShowEditAssignmentModal(true);
                                  }}
                                  style={{ 
                                    display: 'flex', 
                                    alignItems: 'center', 
                                    gap: '8px', 
                                    padding: '14px 0', 
                                    borderBottom: '1px solid var(--border-color)',
                                    cursor: 'pointer',
                                    width: '100%',
                                    boxSizing: 'border-box'
                                  }}
                                  className="hoverable-row-div"
                                  title={language === 'uz' ? "Tafsilotlarni ko'rish" : 'Посмотреть детали'}
                                >
                                  <span style={{ 
                                    fontWeight: '600', 
                                    color: 'var(--accent-color)', 
                                    textDecoration: 'underline',
                                    fontSize: '13.5px',
                                    wordBreak: 'break-word',
                                    lineHeight: '1.3'
                                  }}>
                                    {prod.productName}
                                  </span>
                                  <span style={{ 
                                    fontWeight: '700', 
                                    color: (prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) <= 0 ? 'var(--text-muted)' : 'var(--success-color)',
                                    backgroundColor: (prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) <= 0 ? 'var(--bg-tertiary)' : 'var(--success-light)',
                                    padding: '2px 8px',
                                    borderRadius: '6px',
                                    fontSize: '12px',
                                    whiteSpace: 'nowrap',
                                    flexShrink: 0
                                  }}>
                                    {(prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) <= 0 
                                      ? (language === 'uz' ? 'Sotib bo\'lingan' : 'Продано') 
                                      : `${prod.remainingQty !== undefined ? prod.remainingQty : prod.qty} ${language === 'uz' ? 'dona' : 'шт'}`}
                                  </span>
                                </div>
                              ))}
                            </div>
                          ) : (
                            /* Admin View: Original full table */
                            <table style={{ width: '100%', minWidth: '600px', borderCollapse: 'collapse', textAlign: 'left' }}>
                              <thead>
                                <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '12px' }}>
                                  <th style={{ padding: '8px 4px' }}>{language === 'uz' ? 'Mahsulot' : 'Товар'}</th>
                                  {userRole !== 'agent' && <th style={{ padding: '8px 4px', textAlign: 'right' }}>{language === 'uz' ? 'Olgan' : 'Получил'}</th>}
                                  <th style={{ padding: '8px 4px', textAlign: 'right' }}>{language === 'uz' ? 'Qolgan (Qoldiq)' : 'Остаток'}</th>
                                  {userRole !== 'agent' && <th style={{ padding: '8px 4px', textAlign: 'right' }}>{language === 'uz' ? 'Sana' : 'Дата'}</th>}
                                  {userRole !== 'agent' && <th style={{ padding: '8px 4px', textAlign: 'right' }}>{language === 'uz' ? "O'chirish" : 'Удалить'}</th>}
                                </tr>
                              </thead>
                              <tbody>
                                {agentProducts.map(prod => (
                                  <tr key={prod.id} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '13px' }}>
                                    <td 
                                      onClick={() => {
                                        setSelectedAssignmentForEdit(prod);
                                        setEditAssignmentQty(prod.qty.toString());
                                        setEditAssignmentRemainingQty((prod.remainingQty !== undefined ? prod.remainingQty : prod.qty).toString());
                                        setShowEditAssignmentModal(true);
                                      }}
                                      style={{ 
                                        padding: '10px 4px', 
                                        fontWeight: '600', 
                                        color: 'var(--accent-color)', 
                                        cursor: 'pointer',
                                        textDecoration: 'underline'
                                      }}
                                      title={userRole === 'agent' ? (language === 'uz' ? "Tafsilotlarni ko'rish" : 'Посмотреть детали') : (language === 'uz' ? 'Tahrirlash uchun bosing' : 'Нажмите для редактирования')}
                                    >
                                      {prod.productName}
                                    </td>
                                    {userRole !== 'agent' && <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600' }}>{prod.qty} {language === 'uz' ? 'dona' : 'шт'}</td>}
                                    <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '700', color: (prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) <= 0 ? 'var(--text-muted)' : 'var(--success-color)' }}>
                                      {(prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) <= 0 ? (language === 'uz' ? 'Sotib bo\'lingan' : 'Продано') : (prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) + ' ' + (language === 'uz' ? 'dona' : 'шт')}
                                    </td>
                                    {userRole !== 'agent' && <td style={{ padding: '10px 4px', textAlign: 'right', color: 'var(--text-secondary)' }}>{prod.date}</td>}
                                    {userRole !== 'agent' && (
                                      <td style={{ padding: '10px 4px', textAlign: 'right' }}>
                                        <button
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            handleDeleteProductAssignment(prod.id);
                                          }}
                                          title={language === 'uz' ? "O'chirish" : "Удалить"}
                                          style={{
                                            border: 'none',
                                            backgroundColor: 'transparent',
                                            color: 'var(--warning-color)',
                                            padding: '4px',
                                            cursor: 'pointer',
                                            borderRadius: '4px',
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            justifyContent: 'center'
                                          }}
                                        >
                                          <Trash2 size={15} />
                                        </button>
                                      </td>
                                    )}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </div>
                      </div>


                    </div>
                  </div>
                );
              })()}
            </div>
          )}

          {/* VIEW: AGENT HISTORY */}
          {activeTab === 'agent_history' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '700', color: 'var(--text-primary)', marginBottom: '8px' }}>
                  {language === 'uz' ? "Sotuvlar va Tashriflar Tarixi" : "История продаж и визитов"}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '20px' }}>
                  {language === 'uz' ? "Bugun tashrif buyurilgan va savdo qilingan do'konlar ro'yxati" : "Список магазинов, которые вы посетили сегодня"}
                </p>

                {visitedStores.filter(v => v.date === new Date().toISOString().split('T')[0]).length === 0 ? (
                  <div style={{ padding: '36px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13px' }}>
                    {language === 'uz' ? "Bugun hali hech qaysi do'konga tashrif buyurilmadi." : "Сегодня вы еще не посетили ни один магазин."}
                  </div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    {visitedStores.filter(v => v.date === new Date().toISOString().split('T')[0]).map((visit, index) => (
                      <div key={index} style={{
                        backgroundColor: 'var(--bg-primary)',
                        border: '1px solid var(--border-color)',
                        borderRadius: '8px',
                        padding: '16px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '12px'
                      }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                          <div>
                            <span style={{ fontWeight: '600', color: 'var(--text-primary)', fontSize: '14px' }}>
                              {visit.storeName}
                            </span>
                            <span style={{ fontSize: '11px', color: 'var(--text-muted)', marginLeft: '8px' }}>
                              🕒 {visit.time}
                            </span>
                          </div>
                          <div>
                            {visit.status === 'sold' ? (
                              <span style={{
                                backgroundColor: 'rgba(16, 185, 129, 0.1)',
                                color: '#10b981',
                                padding: '4px 10px',
                                borderRadius: '20px',
                                fontSize: '11px',
                                fontWeight: '600'
                              }}>
                                {language === 'uz' ? "Sotuv yakunlandi" : "Продажа завершена"}
                              </span>
                            ) : (
                              <span style={{
                                backgroundColor: 'rgba(245, 158, 11, 0.1)',
                                color: '#f59e0b',
                                padding: '4px 10px',
                                borderRadius: '20px',
                                fontSize: '11px',
                                fontWeight: '600'
                              }}>
                                {language === 'uz' ? `Sotuvsiz (Sabab: ${visit.reason})` : `Без продажи (Причина: ${visit.reason})`}
                              </span>
                            )}
                          </div>
                        </div>

                        {visit.status === 'sold' && visit.items && visit.items.length > 0 && (
                          <div style={{
                            borderTop: '1px dashed var(--border-color)',
                            paddingTop: '8px',
                            display: 'flex',
                            flexDirection: 'column',
                            gap: '6px'
                          }}>
                            <span style={{ fontSize: '11px', color: 'var(--text-secondary)', fontWeight: '600' }}>
                              {language === 'uz' ? "Sotilgan mahsulotlar:" : "Проданные товары:"}
                            </span>
                            {visit.items.map((item, itemIdx) => (
                              <div key={itemIdx} style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', color: 'var(--text-secondary)' }}>
                                <span>• {item.productName} ({item.qty} dona)</span>
                                <span>{(item.qty * item.price).toLocaleString()} UZS</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* VIEW 5: SALES HISTORY */}
          {activeTab === 'sales' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              
              {/* Date Filter selector */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap', alignSelf: 'flex-start' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '8px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <span style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: '500' }}>{t('date_filter')}:</span>
                  <select 
                    value={dateFilter} 
                    onChange={(e) => setDateFilter(e.target.value)}
                    style={{
                      backgroundColor: 'transparent',
                      border: 'none',
                      color: 'var(--accent-color)',
                      fontSize: '13px',
                      fontWeight: '600',
                      cursor: 'pointer',
                      outline: 'none'
                    }}
                  >
                    <option value="bugun">{t('bugun')}</option>
                    <option value="custom">{t('custom_range')}</option>
                  </select>
                </div>

                {dateFilter === 'custom' && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '6px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', animation: 'fadeIn 0.2s ease' }}>
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('dan')}</span>
                    <input 
                      type="date" 
                      value={startDate} 
                      onChange={(e) => setStartDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('gacha')}</span>
                    <input 
                      type="date" 
                      value={endDate} 
                      onChange={(e) => setEndDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                  </div>
                )}
              </div>

              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '20px' }}>{language === 'uz' ? 'Savdolar va toʻlovlar' : 'Продажи и платежи'}</h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Savdo ID' : 'ID Продажи'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Agent' : 'Агент'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Do\'kon' : 'Магазин'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Summa' : 'Сумма'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'To\'lov turi' : 'Тип оплаты'}</th>
                      <th style={{ padding: '12px' }}>{t('date')}</th>
                      <th style={{ padding: '12px' }}>{t('status')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {getFilteredSales(sales).map((sale) => (
                      <tr key={sale.id} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}>
                        <td style={{ padding: '12px', color: 'var(--accent-color)', fontWeight: '600' }}>#TR-{sale.id}</td>
                        <td style={{ padding: '12px', fontWeight: '500' }}>{getAgentLoginByName(sale.agent)}</td>
                        <td style={{ padding: '12px' }}>{sale.store}</td>
                        <td style={{ padding: '12px', fontWeight: '600' }}>{sale.total.toLocaleString()} UZS</td>
                        <td style={{ padding: '12px' }}>
                          <span style={{
                            fontSize: '11px',
                            fontWeight: '600',
                            backgroundColor: sale.payment === 'Payme' ? 'rgba(13, 148, 136, 0.15)' : 'rgba(99, 102, 241, 0.15)',
                            color: sale.payment === 'Payme' ? 'var(--accent-color)' : '#6366f1',
                            padding: '3px 8px',
                            borderRadius: '4px'
                          }}>{sale.payment}</span>
                        </td>
                        <td style={{ padding: '12px', color: 'var(--text-secondary)' }}>{sale.date}</td>
                        <td style={{ padding: '12px' }}>
                          <span style={{
                            fontSize: '11px',
                            fontWeight: '600',
                            backgroundColor: 'rgba(16, 185, 129, 0.15)',
                            color: 'var(--success-color)',
                            padding: '3px 8px',
                            borderRadius: '4px',
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '4px'
                          }}>
                            <Check size={12} /> {language === 'uz' ? 'to\'landi' : 'оплачено'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

            </div>
          )}

          {/* VIEW: TAHLIL - UMUMIY */}
          {activeTab === 'tahlil_umumiy' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '32px' }}>
              
              {/* Date Filter selector */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap', alignSelf: 'flex-start' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '8px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <span style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: '500' }}>{t('date_filter')}:</span>
                  <select 
                    value={dateFilter} 
                    onChange={(e) => setDateFilter(e.target.value)}
                    style={{
                      backgroundColor: 'transparent',
                      border: 'none',
                      color: 'var(--accent-color)',
                      fontSize: '13px',
                      fontWeight: '600',
                      cursor: 'pointer',
                      outline: 'none'
                    }}
                  >
                    <option value="bugun">{t('bugun')}</option>
                    <option value="custom">{t('custom_range')}</option>
                  </select>
                </div>

                {dateFilter === 'custom' && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '6px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', animation: 'fadeIn 0.2s ease' }}>
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('dan')}</span>
                    <input 
                      type="date" 
                      value={startDate} 
                      onChange={(e) => setStartDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('gacha')}</span>
                    <input 
                      type="date" 
                      value={endDate} 
                      onChange={(e) => setEndDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                  </div>
                )}
              </div>

              {/* Summary Cards */}
              {(() => {
                const filteredSales = getFilteredSales(sales);
                let totalRevenue = 0;
                let totalCost = 0;
                
                filteredSales.forEach(s => {
                  totalRevenue += s.total;
                  if (s.items) {
                    s.items.forEach(item => {
                      const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
                      totalCost += (cost * item.qty);
                    });
                  }
                });

                const totalProfit = totalRevenue - totalCost;

                // Top Agent by sales
                const agentSales = {};
                filteredSales.forEach(s => {
                  agentSales[s.agent] = (agentSales[s.agent] || 0) + s.total;
                });
                let topAgentName = "-";
                let topAgentVal = 0;
                Object.keys(agentSales).forEach(name => {
                  if (agentSales[name] > topAgentVal) {
                    topAgentVal = agentSales[name];
                    topAgentName = name;
                  }
                });

                // Top Store by sales
                const storeSales = {};
                filteredSales.forEach(s => {
                  storeSales[s.store] = (storeSales[s.store] || 0) + s.total;
                });
                let topStoreName = "-";
                let topStoreVal = 0;
                Object.keys(storeSales).forEach(name => {
                  if (storeSales[name] > topStoreVal) {
                    topStoreVal = storeSales[name];
                    topStoreName = name;
                  }
                });

                // Top Product by quantity
                const productQty = {};
                filteredSales.forEach(s => {
                  if (s.items) {
                    s.items.forEach(item => {
                      productQty[item.productName] = (productQty[item.productName] || 0) + item.qty;
                    });
                  }
                });
                let topProductName = "-";
                let topProductVal = 0;
                Object.keys(productQty).forEach(name => {
                  if (productQty[name] > topProductVal) {
                    topProductVal = productQty[name];
                    topProductName = name;
                  }
                });

                return (
                  <>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '20px' }}>
                      
                      {/* Card 1 */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Umumiy Savdo Tushumi' : 'Общая выручка от продаж'}</span>
                        <h3 style={{ fontSize: '24px', fontWeight: '700', color: 'var(--accent-color)', margin: 0 }}>{totalRevenue.toLocaleString()} UZS</h3>
                        <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Jami amalga oshirilgan savdolar' : 'Всего осуществлено продаж'}</span>
                      </div>

                      {/* Card 2 */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Jami Tannarx (Xarajat)' : 'Общая себестоимость (Расход)'}</span>
                        <h3 style={{ fontSize: '24px', fontWeight: '700', color: 'var(--warning-color)', margin: 0 }}>{totalCost.toLocaleString()} UZS</h3>
                        <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Mahsulotlar asl narxi yig\'indisi' : 'Сумма себестоимости товаров'}</span>
                      </div>

                      {/* Card 3 */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{t('net_profit')}</span>
                        <h3 style={{ fontSize: '24px', fontWeight: '700', color: 'var(--success-color)', margin: 0 }}>{totalProfit.toLocaleString()} UZS</h3>
                        <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Tushumdan tannarx ayirilganda' : 'Выручка за вычетом себестоимости'}</span>
                      </div>

                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '20px' }}>
                      
                      {/* Top Agent */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '20px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '50%', backgroundColor: 'rgba(13, 148, 136, 0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent-color)' }}>
                          <UserCheck size={24} />
                        </div>
                        <div>
                          <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('top_agent')}</span>
                          <h4 style={{ fontSize: '15px', fontWeight: '600', margin: '2px 0 0 0' }}>{getAgentLoginByName(topAgentName)}</h4>
                          <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Savdo' : 'Продажи'}: {topAgentVal.toLocaleString()} UZS</span>
                        </div>
                      </div>

                      {/* Top Store */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '20px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '50%', backgroundColor: 'rgba(99, 102, 241, 0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#6366f1' }}>
                          <StoreIcon size={24} />
                        </div>
                        <div>
                          <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('top_store')}</span>
                          <h4 style={{ fontSize: '15px', fontWeight: '600', margin: '2px 0 0 0' }}>{topStoreName}</h4>
                          <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Xarid' : 'Покупка'}: {topStoreVal.toLocaleString()} UZS</span>
                        </div>
                      </div>

                      {/* Top Product */}
                      <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '20px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', alignItems: 'center', gap: '16px' }}>
                        <div style={{ width: '48px', height: '48px', borderRadius: '50%', backgroundColor: 'rgba(16, 185, 129, 0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--success-color)' }}>
                          <Package size={24} />
                        </div>
                        <div>
                          <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('top_product')}</span>
                          <h4 style={{ fontSize: '15px', fontWeight: '600', margin: '2px 0 0 0' }}>{topProductName}</h4>
                          <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>{language === 'uz' ? 'Sotildi' : 'Продано'}: {topProductVal.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</span>
                        </div>
                      </div>

                    </div>
                  </>
                );
              })()}

              {/* Chart section */}
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '20px' }}>{language === 'uz' ? 'Savdo va Sof Foyda Solishtirmasi' : 'Сравнение продаж и чистой прибыли'}</h3>
                <div style={{ width: '100%', height: '300px' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart
                      data={(() => {
                        const dailyData = {};
                        getFilteredSales(sales).forEach(s => {
                          const date = s.date;
                          if (!dailyData[date]) dailyData[date] = { name: date, tushum: 0, foyda: 0 };
                          dailyData[date].tushum += s.total;
                          let saleCost = 0;
                          if (s.items) {
                            s.items.forEach(item => {
                              const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
                              saleCost += (cost * item.qty);
                            });
                          }
                          dailyData[date].foyda += (s.total - saleCost);
                        });
                        return Object.values(dailyData);
                      })()}
                      margin={{ top: 10, right: 30, left: 0, bottom: 0 }}
                    >
                      <defs>
                        <linearGradient id="colorTushum" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="var(--accent-color)" stopOpacity={0.4}/>
                          <stop offset="95%" stopColor="var(--accent-color)" stopOpacity={0}/>
                        </linearGradient>
                        <linearGradient id="colorFoyda" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="var(--success-color)" stopOpacity={0.4}/>
                          <stop offset="95%" stopColor="var(--success-color)" stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="var(--border-color)" />
                      <XAxis dataKey="name" stroke="var(--text-secondary)" style={{ fontSize: '12px' }} />
                      <YAxis stroke="var(--text-secondary)" style={{ fontSize: '12px' }} />
                      <Tooltip contentStyle={{ backgroundColor: 'var(--bg-secondary)', border: '1px solid var(--border-color)', color: 'var(--text-primary)' }} />
                      <Area type="monotone" dataKey="tushum" name={language === 'uz' ? 'Tushum' : 'Выручка'} stroke="var(--accent-color)" fillOpacity={1} fill="url(#colorTushum)" />
                      <Area type="monotone" dataKey="foyda" name={language === 'uz' ? 'Sof Foyda' : 'Чистая прибыль'} stroke="var(--success-color)" fillOpacity={1} fill="url(#colorFoyda)" />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>

            </div>
          )}

          {/* VIEW: TAHLIL - DO'KON */}
          {activeTab === 'tahlil_dokon' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              
              {/* Filter controls */}
              <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap', width: '100%' }}>
                <div style={{ position: 'relative', flexGrow: 1 }}>
                  <span style={{ position: 'absolute', left: '14px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-secondary)', display: 'flex', alignItems: 'center' }}>
                    <Search size={16} />
                  </span>
                  <input 
                    type="text" 
                    placeholder={language === 'uz' ? 'Do\'kon nomini qidirish...' : 'Поиск магазина по названию...'}
                    value={productSearchQuery} 
                    onChange={(e) => setProductSearchQuery(e.target.value)}
                    style={{
                      width: '100%',
                      padding: '11px 16px 11px 40px',
                      borderRadius: '8px',
                      border: '1px solid var(--border-color)',
                      backgroundColor: 'var(--bg-secondary)',
                      color: 'var(--text-primary)',
                      fontSize: '14px',
                      outline: 'none',
                      transition: 'border-color var(--transition-fast)'
                    }}
                  />
                </div>

                {/* Date Filter selector */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '11px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: '500' }}>{t('date_filter')}:</span>
                    <select 
                      value={dateFilter} 
                      onChange={(e) => setDateFilter(e.target.value)}
                      style={{
                        backgroundColor: 'transparent',
                        border: 'none',
                        color: 'var(--accent-color)',
                        fontSize: '13px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        outline: 'none'
                      }}
                    >
                      <option value="bugun">{t('bugun')}</option>
                      <option value="custom">{t('custom_range')}</option>
                    </select>
                  </div>

                  {dateFilter === 'custom' && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '9px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', animation: 'fadeIn 0.2s ease' }}>
                      <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('dan')}</span>
                      <input 
                        type="date" 
                        value={startDate} 
                        onChange={(e) => setStartDate(e.target.value)} 
                        style={{
                          border: 'none',
                          backgroundColor: 'transparent',
                          color: 'var(--text-primary)',
                          fontSize: '13px',
                          fontWeight: '600',
                          outline: 'none',
                          cursor: 'pointer'
                        }}
                      />
                      <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('gacha')}</span>
                      <input 
                        type="date" 
                        value={endDate} 
                        onChange={(e) => setEndDate(e.target.value)} 
                        style={{
                          border: 'none',
                          backgroundColor: 'transparent',
                          color: 'var(--text-primary)',
                          fontSize: '13px',
                          fontWeight: '600',
                          outline: 'none',
                          cursor: 'pointer'
                        }}
                      />
                    </div>
                  )}
                </div>
              </div>

              {/* Table of Stores */}
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Do\'kon nomi' : 'Название магазина'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Xarid qilingan mahsulotlar (Soni)' : 'Купленные товары (Кол-во)'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Umumiy tushum' : 'Общая выручка'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Sof foyda' : 'Чистая прибыль'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Eng ko\'p olingan mahsulot' : 'Самый покупаемый товар'}</th>
                      <th style={{ padding: '12px' }}>{language === 'uz' ? 'Sotgan agentlar' : 'Продавшие агенты'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {stores
                      .filter(s => s.name.toLowerCase().includes(productSearchQuery.toLowerCase()))
                      .map(store => {
                        const storeSalesList = getFilteredSales(sales).filter(s => s.storeId === store.id || s.store === store.name);
                        
                        let totalQty = 0;
                        let revenue = 0;
                        let costSum = 0;
                        const productCounts = {};
                        const agentsSet = new Set();

                        storeSalesList.forEach(s => {
                          revenue += s.total;
                          agentsSet.add(s.agent);
                          if (s.items) {
                            s.items.forEach(item => {
                              totalQty += item.qty;
                              const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
                              costSum += (cost * item.qty);
                              productCounts[item.productName] = (productCounts[item.productName] || 0) + item.qty;
                            });
                          }
                        });

                        const profit = revenue - costSum;

                        // Find top product in store
                        let topProd = "-";
                        let topProdQty = 0;
                        Object.keys(productCounts).forEach(name => {
                          if (productCounts[name] > topProdQty) {
                            topProdQty = productCounts[name];
                            topProd = name;
                          }
                        });

                        return (
                          <tr 
                            key={store.id} 
                            onClick={() => setSelectedTahlilStoreDetails(store)} 
                            className="hoverable-row"
                            style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}
                          >
                            <td style={{ padding: '12px', fontWeight: '600' }}>
                              <div>{store.name}</div>
                              <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>ID: {store.id}</span>
                            </td>
                            <td style={{ padding: '12px' }}>{totalQty.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</td>
                            <td style={{ padding: '12px', fontWeight: '600', color: 'var(--accent-color)' }}>{revenue.toLocaleString()} UZS</td>
                            <td style={{ padding: '12px', fontWeight: '600', color: profit >= 0 ? 'var(--success-color)' : 'var(--warning-color)' }}>
                              {profit.toLocaleString()} UZS
                            </td>
                            <td style={{ padding: '12px', fontSize: '13px' }}>
                              {topProd} {topProdQty > 0 && `(${topProdQty} ${language === 'uz' ? 'dona' : 'шт'})`}
                            </td>
                            <td style={{ padding: '12px', fontSize: '13px', color: 'var(--text-secondary)' }}>
                              {Array.from(agentsSet).map(name => getAgentLoginByName(name)).join(', ') || (language === 'uz' ? 'Yo\'q' : 'Нет')}
                            </td>
                          </tr>
                        );
                      })}
                  </tbody>
                </table>
              </div>

              {/* Store Sales Details Modal */}
              {selectedTahlilStoreDetails && (
                <div style={{
                  position: 'fixed',
                  top: 0,
                  left: 0,
                  width: '100vw',
                  height: '100vh',
                  backgroundColor: 'rgba(15, 23, 42, 0.75)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 1000
                }} className="fade-in">
                  <div style={{
                    width: '680px',
                    backgroundColor: 'var(--bg-secondary)',
                    border: '1px solid var(--border-color)',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: 'var(--shadow-lg)',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '20px'
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div>
                        <h3 style={{ fontSize: '18px', fontWeight: '600' }}>{selectedTahlilStoreDetails.name}</h3>
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Mijoz do\'kon xaridlari tahlili' : 'Анализ покупок магазина'}</span>
                      </div>
                      <button 
                        onClick={() => setSelectedTahlilStoreDetails(null)}
                        style={{
                          border: 'none',
                          backgroundColor: 'transparent',
                          color: 'var(--text-secondary)',
                          cursor: 'pointer',
                          fontSize: '20px',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center'
                        }}
                      >
                        &times;
                      </button>
                    </div>

                    <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                        <thead>
                          <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                            <th style={{ padding: '10px' }}>{language === 'uz' ? 'Mahsulot nomi' : 'Название товара'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Sotilgan miqdor' : 'Проданное количество'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Jami tushum' : 'Общая выручка'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Sof foyda' : 'Чистая прибыль'}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {(() => {
                            const storeSalesList = getFilteredSales(sales).filter(s => s.storeId === selectedTahlilStoreDetails.id || s.store === selectedTahlilStoreDetails.name);
                            const productsMap = {};

                            storeSalesList.forEach(s => {
                              if (s.items) {
                                s.items.forEach(item => {
                                  if (!productsMap[item.productName]) {
                                    productsMap[item.productName] = { name: item.productName, qty: 0, revenue: 0, cost: 0 };
                                  }
                                  productsMap[item.productName].qty += item.qty;
                                  productsMap[item.productName].revenue += (item.qty * item.price);
                                  const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
                                  productsMap[item.productName].cost += (item.qty * cost);
                                });
                              }
                            });

                            const sortedProducts = Object.values(productsMap).sort((a, b) => b.revenue - a.revenue);

                            if (sortedProducts.length === 0) {
                              return (
                                <tr>
                                  <td colSpan="4" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                                    {language === 'uz' ? 'Ushbu do\'konga hali mahsulot sotilmagan.' : 'В этот магазин еще не было продаж.'}
                                  </td>
                                </tr>
                              );
                            }

                            return sortedProducts.map(item => {
                              const itemProfit = item.revenue - item.cost;
                              return (
                                <tr key={item.name} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}>
                                  <td style={{ padding: '10px', fontWeight: '500' }}>{item.name}</td>
                                  <td style={{ padding: '10px', textAlign: 'right' }}>{item.qty.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</td>
                                  <td style={{ padding: '10px', textAlign: 'right', fontWeight: '600', color: 'var(--accent-color)' }}>{item.revenue.toLocaleString()} UZS</td>
                                  <td style={{ padding: '10px', textAlign: 'right', fontWeight: '600', color: itemProfit >= 0 ? 'var(--success-color)' : 'var(--warning-color)' }}>
                                    {itemProfit.toLocaleString()} UZS
                                  </td>
                                </tr>
                              );
                            });
                          })()}
                        </tbody>
                      </table>
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '10px' }}>
                      <button 
                        onClick={() => setSelectedTahlilStoreDetails(null)}
                        style={{
                          padding: '10px 20px',
                          backgroundColor: 'var(--bg-tertiary)',
                          color: 'var(--text-primary)',
                          border: 'none',
                          borderRadius: '8px',
                          cursor: 'pointer',
                          fontWeight: '600',
                          fontSize: '14px'
                        }}
                      >
                        {t('close')}
                      </button>
                    </div>

                  </div>
                </div>
              )}

            </div>
          )}

          {/* VIEW: TAHLIL - AGENT */}
          {activeTab === 'tahlil_agent' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px', width: '100%' }}>
              
              {/* Date Filter selector */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap', alignSelf: 'flex-start' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '8px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <span style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: '500' }}>{t('date_filter')}:</span>
                  <select 
                    value={dateFilter} 
                    onChange={(e) => setDateFilter(e.target.value)}
                    style={{
                      backgroundColor: 'transparent',
                      border: 'none',
                      color: 'var(--accent-color)',
                      fontSize: '13px',
                      fontWeight: '600',
                      cursor: 'pointer',
                      outline: 'none'
                    }}
                  >
                    <option value="bugun">{t('bugun')}</option>
                    <option value="custom">{t('custom_range')}</option>
                  </select>
                </div>

                {dateFilter === 'custom' && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: 'var(--bg-secondary)', padding: '6px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', animation: 'fadeIn 0.2s ease' }}>
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('dan')}</span>
                    <input 
                      type="date" 
                      value={startDate} 
                      onChange={(e) => setStartDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{t('gacha')}</span>
                    <input 
                      type="date" 
                      value={endDate} 
                      onChange={(e) => setEndDate(e.target.value)} 
                      style={{
                        border: 'none',
                        backgroundColor: 'transparent',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        fontWeight: '600',
                        outline: 'none',
                        cursor: 'pointer'
                      }}
                    />
                  </div>
                )}
              </div>

              <div style={{ display: 'flex', gap: '32px', flexWrap: 'wrap' }}>
              
              {/* Left Column: Agents List */}
              <div style={{
                flex: '1 1 300px',
                backgroundColor: 'var(--bg-secondary)',
                border: '1px solid var(--border-color)',
                borderRadius: '12px',
                padding: '24px',
                display: 'flex',
                flexDirection: 'column',
                gap: '16px'
              }}>
                <h3 style={{ fontSize: '16px', fontWeight: '600', margin: 0 }}>{t('agents_list')}</h3>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '500px', overflowY: 'auto' }}>
                  {agents.map(agent => {
                    const agentSalesList = getFilteredSales(sales).filter(s => s.agentId === agent.id || s.agent === agent.name);
                    const totalRevenue = agentSalesList.reduce((sum, s) => sum + s.total, 0);

                    return (
                      <div 
                        key={agent.id}
                        onClick={() => setSelectedTahlilAgentId(agent.id)}
                        style={{
                          padding: '12px 16px',
                          borderRadius: '8px',
                          border: '1px solid var(--border-color)',
                          backgroundColor: selectedTahlilAgentId === agent.id ? 'var(--accent-light)' : 'var(--bg-primary)',
                          cursor: 'pointer',
                          transition: 'all var(--transition-fast)'
                        }}
                      >
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <span style={{ fontWeight: '600', color: selectedTahlilAgentId === agent.id ? 'var(--accent-color)' : 'var(--text-primary)' }}>
                            {agent.login}
                          </span>
                          <span style={{ fontSize: '12px', fontWeight: '700', color: 'var(--text-secondary)' }}>
                            {totalRevenue.toLocaleString()} UZS
                          </span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '4px', fontSize: '11px', color: 'var(--text-muted)' }}>
                          <span>{agent.login}</span>
                          <span>@{agent.username || 'username'}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Right Column: Selected Agent Analytics Details */}
              {(() => {
                const selAgent = agents.find(a => a.id === selectedTahlilAgentId);
                if (!selAgent) {
                  return (
                    <div style={{ flex: '2 1 600px', backgroundColor: 'var(--bg-secondary)', border: '1px solid var(--border-color)', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--text-secondary)' }}>
                      {t('select_agent_desc')}
                    </div>
                  );
                }

                const agentSalesList = getFilteredSales(sales).filter(s => s.agentId === selAgent.id || s.agent === selAgent.name);
                
                let totalRevenue = 0;
                let totalCost = 0;
                const storeVisits = new Set();
                const productSummary = {};

                agentSalesList.forEach(s => {
                  totalRevenue += s.total;
                  storeVisits.add(s.store);
                  if (s.items) {
                    s.items.forEach(item => {
                      const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
                      totalCost += (cost * item.qty);
                      
                      if (!productSummary[item.productName]) {
                        productSummary[item.productName] = { qty: 0, revenue: 0, cost: 0 };
                      }
                      productSummary[item.productName].qty += item.qty;
                      productSummary[item.productName].revenue += (item.qty * item.price);
                      productSummary[item.productName].cost += (item.qty * cost);
                    });
                  }
                });

                const totalProfit = totalRevenue - totalCost;

                return (
                  <div style={{ flex: '2 1 600px', display: 'flex', flexDirection: 'column', gap: '24px' }}>
                    
                    {/* Header metrics */}
                    <div style={{
                      backgroundColor: 'var(--bg-secondary)',
                      border: '1px solid var(--border-color)',
                      borderRadius: '12px',
                      padding: '24px',
                      display: 'flex',
                      flexWrap: 'wrap',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                      gap: '16px'
                    }}>
                      <div>
                        <h2 style={{ fontSize: '20px', fontWeight: '700', margin: 0 }}>{selAgent.login} {language === 'uz' ? 'Tahlili' : 'Анализ'}</h2>
                        <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                          Username: @{selAgent.username || 'username'} | Tel: {selAgent.phone || (language === 'uz' ? 'Kiritilmagan' : 'Не указано')}
                        </span>
                      </div>
                      
                      <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap' }}>
                        <div style={{ textAlign: 'right' }}>
                          <div style={{ fontSize: '11px', color: 'var(--text-secondary)', textTransform: 'uppercase' }}>{language === 'uz' ? 'Jami Savdo' : 'Всего продаж'}</div>
                          <div style={{ fontSize: '16px', fontWeight: '700', color: 'var(--accent-color)' }}>{totalRevenue.toLocaleString()} UZS</div>
                        </div>
                        <div style={{ textAlign: 'right' }}>
                          <div style={{ fontSize: '11px', color: 'var(--text-secondary)', textTransform: 'uppercase' }}>{language === 'uz' ? 'Sof Foyda' : 'Чистая прибыль'}</div>
                          <div style={{ fontSize: '16px', fontWeight: '700', color: 'var(--success-color)' }}>{totalProfit.toLocaleString()} UZS</div>
                        </div>
                        <div style={{ textAlign: 'right' }}>
                          <div style={{ fontSize: '11px', color: 'var(--text-secondary)', textTransform: 'uppercase' }}>{language === 'uz' ? 'Do\'konlar' : 'Магазины'}</div>
                          <div style={{ fontSize: '16px', fontWeight: '700', color: 'var(--text-primary)' }}>{storeVisits.size} {language === 'uz' ? 'ta' : 'шт'}</div>
                        </div>
                      </div>
                    </div>

                    {/* Table 1: Sotilgan mahsulotlar batafsil */}
                    <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                      <h3 style={{ fontSize: '15px', fontWeight: '600', marginBottom: '16px' }}>{language === 'uz' ? 'Sotilgan mahsulotlar tafsiloti' : 'Детали проданных товаров'}</h3>
                      <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                        <thead>
                          <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                            <th style={{ padding: '10px' }}>{language === 'uz' ? 'Mahsulot nomi' : 'Название товара'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Sotilgan miqdor' : 'Проданное количество'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Sotuv summasi' : 'Сумма продаж'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Sof foyda' : 'Чистая прибыль'}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {Object.keys(productSummary).length === 0 ? (
                            <tr>
                              <td colSpan="4" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                                {language === 'uz' ? 'Bu agent hozircha savdo qilmagan.' : 'Этот агент еще не совершал продаж.'}
                              </td>
                            </tr>
                          ) : (
                            Object.keys(productSummary).map(prodName => {
                              const info = productSummary[prodName];
                              const pProfit = info.revenue - info.cost;
                              return (
                                <tr 
                                  key={prodName} 
                                  className="hoverable-row"
                                  onClick={() => setSelectedTahlilProductDetails(prodName)}
                                  style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}
                                >
                                  <td style={{ padding: '10px', fontWeight: '500', color: 'var(--accent-color)', textDecoration: 'underline', textDecorationStyle: 'dashed' }}>{prodName}</td>
                                  <td style={{ padding: '10px', textAlign: 'right' }}>{info.qty.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</td>
                                  <td style={{ padding: '10px', textAlign: 'right', fontWeight: '600', color: 'var(--accent-color)' }}>{info.revenue.toLocaleString()} UZS</td>
                                  <td style={{ padding: '10px', textAlign: 'right', fontWeight: '600', color: pProfit >= 0 ? 'var(--success-color)' : 'var(--warning-color)' }}>
                                    {pProfit.toLocaleString()} UZS
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>

                    {/* Table 2: Tashrif buyurilgan do'konlar */}
                    <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)' }}>
                      <h3 style={{ fontSize: '15px', fontWeight: '600', marginBottom: '16px' }}>{language === 'uz' ? 'Savdo qilingan do\'konlar' : 'Магазины, в которых велись продажи'}</h3>
                      <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                        <thead>
                          <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '13px' }}>
                            <th style={{ padding: '10px' }}>{language === 'uz' ? 'Do\'kon nomi' : 'Название магазина'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Xarid summasi' : 'Сумма покупки'}</th>
                            <th style={{ padding: '10px', textAlign: 'right' }}>{language === 'uz' ? 'Oxirgi savdo sanasi' : 'Дата последней продажи'}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {agentSalesList.length === 0 ? (
                            <tr>
                              <td colSpan="3" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                                {language === 'uz' ? 'Ushbu agentga tegishli do\'kon savdolari topilmadi.' : 'Продаж этого агента в магазины не найдено.'}
                              </td>
                            </tr>
                          ) : (
                            agentSalesList.map(sale => (
                              <tr 
                                key={sale.id} 
                                className="hoverable-row"
                                onClick={() => setSelectedTahlilAgentStore(sale.store)}
                                style={{ borderBottom: '1px solid var(--border-color)', fontSize: '14px' }}
                              >
                                <td style={{ padding: '10px', fontWeight: '500', color: 'var(--accent-color)', textDecoration: 'underline', textDecorationStyle: 'dashed' }}>{sale.store}</td>
                                <td style={{ padding: '10px', textAlign: 'right', fontWeight: '600', color: 'var(--accent-color)' }}>{sale.total.toLocaleString()} UZS</td>
                                <td style={{ padding: '10px', textAlign: 'right', color: 'var(--text-secondary)' }}>{sale.date}</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>

                  </div>
                );
              })()}

              </div>
            </div>
          )}
          {/* VIEW: SETTINGS PROFILE */}
          {activeTab === 'settings_profile' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px', maxWidth: '600px' }}>
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '32px', borderRadius: '16px', border: '1px solid var(--border-color)', boxShadow: 'var(--shadow-sm)' }}>
                <h3 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)', marginBottom: '4px' }}>
                  {t('profile_title')}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '24px' }}>
                  {t('profile_desc')}
                </p>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                  <div>
                    <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '8px', fontWeight: '600' }}>
                      {t('admin_name_label')}
                    </label>
                    <input 
                      type="text" 
                      value={adminName}
                      onChange={(e) => setAdminName(e.target.value)}
                      placeholder="Admin"
                      style={{
                        width: '100%',
                        padding: '12px 16px',
                        borderRadius: '8px',
                        border: '1px solid var(--border-color)',
                        backgroundColor: 'var(--bg-primary)',
                        color: 'var(--text-primary)',
                        fontSize: '14px',
                        outline: 'none'
                      }}
                    />
                  </div>

                  <div>
                    <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '8px', fontWeight: '600' }}>
                      {t('admin_photo_label')}
                    </label>
                    
                    <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                      <div style={{
                        width: '80px',
                        height: '80px',
                        borderRadius: '50%',
                        backgroundColor: 'var(--bg-tertiary)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        fontWeight: '600',
                        fontSize: '24px',
                        overflow: 'hidden',
                        border: '2px solid var(--border-color)'
                      }}>
                        {adminPhoto ? (
                          <img src={adminPhoto} style={{ width: '100%', height: '100%', objectFit: 'cover' }} alt="Admin Profile" />
                        ) : (
                          'AD'
                        )}
                      </div>

                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <input 
                          type="file" 
                          id="admin-photo-input" 
                          accept="image/*"
                          onChange={(e) => {
                            const file = e.target.files[0];
                            if (file) {
                              const reader = new FileReader();
                              reader.onload = (evt) => {
                                setAdminPhoto(evt.target.result);
                              };
                              reader.readAsDataURL(file);
                            }
                          }}
                          style={{ display: 'none' }}
                        />
                        <button
                          type="button"
                          onClick={() => document.getElementById('admin-photo-input').click()}
                          style={{
                            padding: '10px 18px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'var(--bg-primary)',
                            color: 'var(--text-primary)',
                            fontSize: '13px',
                            fontWeight: '600',
                            cursor: 'pointer',
                            display: 'inline-flex',
                            alignItems: 'center',
                            gap: '8px'
                          }}
                        >
                          {t('upload_photo_btn')}
                        </button>
                        {adminPhoto && (
                          <button
                            type="button"
                            onClick={() => setAdminPhoto('')}
                            style={{
                              alignSelf: 'flex-start',
                              fontSize: '11px',
                              color: 'var(--danger-color)',
                              background: 'none',
                              border: 'none',
                              cursor: 'pointer',
                              padding: 0
                            }}
                          >
                            {language === 'uz' ? "Rasm o'chirish" : "Удалить фото"}
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => {
                      localStorage.setItem('adminName', adminName);
                      localStorage.setItem('adminPhoto', adminPhoto);
                      showAlert(t('save_success'), 'success');
                    }}
                    style={{
                      padding: '12px 24px',
                      borderRadius: '8px',
                      border: 'none',
                      backgroundColor: 'var(--accent-color)',
                      color: '#fff',
                      fontSize: '14px',
                      fontWeight: '600',
                      cursor: 'pointer',
                      alignSelf: 'flex-start',
                      marginTop: '8px',
                      transition: 'background var(--transition-fast)'
                    }}
                    onMouseOver={(e) => e.target.style.backgroundColor = 'var(--accent-hover)'}
                    onMouseOut={(e) => e.target.style.backgroundColor = 'var(--accent-color)'}
                  >
                    {t('save_btn')}
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* VIEW: SETTINGS LANGUAGE */}
          {activeTab === 'settings_language' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px', maxWidth: '600px' }}>
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '32px', borderRadius: '16px', border: '1px solid var(--border-color)', boxShadow: 'var(--shadow-sm)' }}>
                <h3 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)', marginBottom: '4px' }}>
                  {t('language_title')}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '24px' }}>
                  {t('language_desc')}
                </p>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                  <label style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: '600' }}>
                    {t('system_language')}
                  </label>

                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                    {/* Uzbek Card */}
                    <div 
                      onClick={() => {
                        setLanguage('uz');
                        localStorage.setItem('lang', 'uz');
                      }}
                      style={{
                        padding: '20px',
                        borderRadius: '12px',
                        border: '2px solid ' + (language === 'uz' ? 'var(--accent-color)' : 'var(--border-color)'),
                        backgroundColor: language === 'uz' ? 'rgba(13, 148, 136, 0.05)' : 'var(--bg-primary)',
                        cursor: 'pointer',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '8px',
                        transition: 'all var(--transition-fast)'
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <span style={{ fontSize: '15px', fontWeight: '700', color: 'var(--text-primary)' }}>O'zbekcha</span>
                        {language === 'uz' && <Check size={18} style={{ color: 'var(--accent-color)' }} />}
                      </div>
                      <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Tizim interfeysini o\'zbek tiliga o\'tkazish' : 'Переключить интерфейс системы на узбекский язык'}</span>
                    </div>

                    {/* Russian Card */}
                    <div 
                      onClick={() => {
                        setLanguage('ru');
                        localStorage.setItem('lang', 'ru');
                      }}
                      style={{
                        padding: '20px',
                        borderRadius: '12px',
                        border: '2px solid ' + (language === 'ru' ? 'var(--accent-color)' : 'var(--border-color)'),
                        backgroundColor: language === 'ru' ? 'rgba(13, 148, 136, 0.05)' : 'var(--bg-primary)',
                        cursor: 'pointer',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '8px',
                        transition: 'all var(--transition-fast)'
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <span style={{ fontSize: '15px', fontWeight: '700', color: 'var(--text-primary)' }}>Русский</span>
                        {language === 'ru' && <Check size={18} style={{ color: 'var(--accent-color)' }} />}
                      </div>
                      <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Tizim interfeysini rus tiliga o\'tkazish' : 'Переключить интерфейс системы на русский язык'}</span>
                    </div>
                  </div>

                </div>
              </div>
            </div>
          )}

          {/* VIEW: SETTINGS DISCOUNTS */}
          {activeTab === 'settings_discounts' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px', maxWidth: '600px' }}>
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '32px', borderRadius: '16px', border: '1px solid var(--border-color)', boxShadow: 'var(--shadow-sm)' }}>
                <h3 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)', marginBottom: '4px' }}>
                  {t('discounts_title')}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '28px' }}>
                  {t('discounts_desc')}
                </p>

                {/* Switch Toggle Container */}
                <div style={{ 
                  display: 'flex', 
                  alignItems: 'center', 
                  justifyContent: 'space-between', 
                  padding: '20px', 
                  borderRadius: '12px', 
                  backgroundColor: 'var(--bg-primary)', 
                  border: '1px solid var(--border-color)',
                  marginBottom: '32px',
                  gap: '16px'
                }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', flex: 1 }}>
                    <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>
                      {t('custom_discount_toggle_label')}
                    </span>
                    <span style={{ fontSize: '12px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                      {t('custom_discount_toggle_desc')}
                    </span>
                  </div>
                  
                  {/* Switch Toggle */}
                  <div 
                    onClick={() => {
                      const newValue = !customDiscountEnabled;
                      setCustomDiscountEnabled(newValue);
                      localStorage.setItem('customDiscountEnabled', String(newValue));
                      showAlert(
                        language === 'uz'
                          ? (newValue ? "Erkin chegirma berish ruxsat etildi" : "Erkin chegirma berish o'chirildi")
                          : (newValue ? "Произвольная скидка разрешена" : "Произвольная скидка отключена"),
                        'info'
                      );
                    }}
                    style={{
                      width: '52px',
                      height: '28px',
                      borderRadius: '14px',
                      backgroundColor: customDiscountEnabled ? 'var(--accent-color)' : 'rgba(255,255,255,0.1)',
                      position: 'relative',
                      cursor: 'pointer',
                      transition: 'background-color 0.3s ease',
                      flexShrink: 0,
                      border: '1px solid rgba(255,255,255,0.05)'
                    }}
                  >
                    <div style={{
                      width: '22px',
                      height: '22px',
                      borderRadius: '50%',
                      backgroundColor: '#fff',
                      position: 'absolute',
                      top: '2px',
                      left: customDiscountEnabled ? '26px' : '2px',
                      transition: 'left 0.2s cubic-bezier(0.4, 0, 0.2, 1)',
                      boxShadow: '0 2px 4px rgba(0,0,0,0.2)'
                    }} />
                  </div>
                </div>

                {/* Defined Discounts section */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                  <h4 style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>
                    {t('defined_discounts_title')}
                  </h4>

                  {/* Add discount percent form */}
                  <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                    <div style={{ position: 'relative', flex: 1 }}>
                      <input 
                        id="new-discount-input"
                        type="number"
                        min="1"
                        max="99"
                        placeholder={t('enter_discount_placeholder')}
                        style={{
                          width: '100%',
                          padding: '12px 36px 12px 16px',
                          borderRadius: '8px',
                          border: '1px solid var(--border-color)',
                          backgroundColor: 'var(--bg-primary)',
                          color: 'var(--text-primary)',
                          fontSize: '14px',
                          outline: 'none',
                          fontWeight: '500'
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            const val = parseInt(e.target.value);
                            if (val) handleAddDiscount(val);
                          }
                        }}
                      />
                      <span style={{
                        position: 'absolute',
                        right: '16px',
                        top: '50%',
                        transform: 'translateY(-50%)',
                        color: 'var(--text-secondary)',
                        fontSize: '14px',
                        fontWeight: '600'
                      }}>%</span>
                    </div>

                    <button
                      onClick={() => {
                        const input = document.getElementById('new-discount-input');
                        if (input) {
                          const val = parseInt(input.value);
                          if (val) handleAddDiscount(val);
                        }
                      }}
                      style={{
                        padding: '12px 20px',
                        borderRadius: '8px',
                        border: 'none',
                        backgroundColor: 'var(--accent-color)',
                        color: '#fff',
                        fontSize: '14px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        transition: 'background var(--transition-fast)'
                      }}
                      onMouseOver={(e) => e.target.style.backgroundColor = 'var(--accent-hover)'}
                      onMouseOut={(e) => e.target.style.backgroundColor = 'var(--accent-color)'}
                    >
                      {t('add_discount_btn')}
                    </button>
                  </div>

                  {/* List of discounts */}
                  <div style={{ 
                    display: 'flex', 
                    flexWrap: 'wrap', 
                    gap: '12px', 
                    marginTop: '12px'
                  }}>
                    {discountsList.map(disc => (
                      <div 
                        key={disc}
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '8px',
                          padding: '8px 16px',
                          borderRadius: '20px',
                          backgroundColor: 'rgba(13, 148, 136, 0.1)',
                          border: '1px solid rgba(13, 148, 136, 0.3)',
                          color: 'var(--accent-color)',
                          fontSize: '14px',
                          fontWeight: '700'
                        }}
                      >
                        <span>{disc}%</span>
                        <span 
                          onClick={() => handleDeleteDiscount(disc)}
                          style={{ 
                            cursor: 'pointer', 
                            fontSize: '16px', 
                            color: 'var(--warning-color)', 
                            fontWeight: 'bold',
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            marginLeft: '4px'
                          }}
                          title={language === 'uz' ? "O'chirish" : "Удалить"}
                        >
                          &times;
                        </span>
                      </div>
                    ))}
                  </div>

                </div>
              </div>
            </div>
          )}

          {/* VIEW: PAYMENT INTEGRATION SETTINGS */}
          {activeTab === 'settings_payments' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', maxWidth: '600px' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '8px', color: 'var(--text-primary)' }}>
                  {language === 'uz' ? "Tinda Terminal Sozlamalari" : "Настройки терминала Tinda"}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '24px' }}>
                  {language === 'uz' 
                    ? "Kassada to'lovni amalga oshirishda foydalaniladigan Tinda smart-terminalining ulanish parametrlarini sozlang." 
                    : "Настройте параметры подключения смарт-терминала Tinda, используемого для оплаты на кассе."}
                </p>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                  
                  {/* IP Address */}
                  <div>
                    <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                      {language === 'uz' ? "Terminal IP manzili va Port (Host:Port)" : "IP-адрес и порт терминала (Host:Port)"}
                    </label>
                    <input 
                      type="text" 
                      value={tindaTerminalIp}
                      onChange={(e) => setTindaTerminalIp(e.target.value)}
                      placeholder="Masalan: 192.168.1.100:8080"
                      style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                    />
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                    {/* User Name */}
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Terminal Xodim Logini" : "Логин сотрудника на терминале"}
                      </label>
                      <input 
                        type="text" 
                        value={tindaTerminalLogin}
                        onChange={(e) => setTindaTerminalLogin(e.target.value)}
                        placeholder="Masalan: Test Administrator"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                      />
                    </div>

                    {/* PIN Code */}
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Terminal PIN kodi" : "ПИН-код терминала"}
                      </label>
                      <input 
                        type="password" 
                        value={tindaTerminalPin}
                        onChange={(e) => setTindaTerminalPin(e.target.value)}
                        placeholder="Masalan: 1111"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                      />
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '12px' }}>
                    {/* Default MXIK */}
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Birlamchi MXIK (IKPU) kodi" : "Основной код ИКПУ (МХИК)"}
                      </label>
                      <input 
                        type="text" 
                        value={tindaDefaultMxik}
                        onChange={(e) => setTindaDefaultMxik(e.target.value)}
                        placeholder="Masalan: 09901001001000000"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                      />
                    </div>

                    {/* Default Package Code */}
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Birlamchi Qadoq kodi" : "Основной код упаковки"}
                      </label>
                      <input 
                        type="text" 
                        value={tindaDefaultPackage}
                        onChange={(e) => setTindaDefaultPackage(e.target.value)}
                        placeholder="Masalan: 242030"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                      />
                    </div>
                  </div>

                  {/* Save Button */}
                  <button
                    onClick={() => {
                      localStorage.setItem('tinda_terminal_ip', tindaTerminalIp);
                      localStorage.setItem('tinda_terminal_login', tindaTerminalLogin);
                      localStorage.setItem('tinda_terminal_pin', tindaTerminalPin);
                      localStorage.setItem('tinda_default_mxik', tindaDefaultMxik);
                      localStorage.setItem('tinda_default_package', tindaDefaultPackage);
                      showAlert(language === 'uz' ? "Tinda terminal sozlamalari saqlandi!" : "Настройки терминала Tinda сохранены!", 'success');
                    }}
                    style={{
                      padding: '12px',
                      backgroundColor: 'var(--accent-color)',
                      color: '#fff',
                      border: 'none',
                      borderRadius: '8px',
                      fontWeight: '600',
                      fontSize: '14px',
                      cursor: 'pointer',
                      width: '100%',
                      marginTop: '10px'
                    }}
                  >
                    {language === 'uz' ? "Sozlamalarni saqlash" : "Сохранить настройки"}
                  </button>

                </div>
              </div>
            </div>
          )}

          {/* VIEW: COMPANY LOGO & NAME SETTINGS */}
          {activeTab === 'settings_company' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', maxWidth: '600px' }}>
                <h3 style={{ fontSize: '16px', fontWeight: '600', marginBottom: '8px', color: 'var(--text-primary)' }}>
                  {language === 'uz' ? "Firma sozlamalari" : "Настройки фирмы"}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '24px' }}>
                  {language === 'uz' 
                    ? "Sidebar panelidagi logotip, firma nomi, sarlavha va ranglarini sozlang." 
                    : "Настройте логотип, название фирмы, подзаголовок и их цвета в боковой панели."}
                </p>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                  
                  {/* Logo Upload */}
                  <div style={{ borderBottom: '1px solid var(--border-color)', paddingBottom: '20px' }}>
                    <label style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-primary)', display: 'block', marginBottom: '8px' }}>
                      {language === 'uz' ? "Firma Logotipi" : "Логотип фирмы"}
                    </label>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
                      <div style={{ 
                        width: '64px', 
                        height: '64px', 
                        borderRadius: '12px', 
                        border: '2px dashed var(--border-light)', 
                        backgroundColor: 'var(--bg-primary)', 
                        display: 'flex', 
                        alignItems: 'center', 
                        justifyContent: 'center',
                        overflow: 'hidden'
                      }}>
                        {companyLogo ? (
                          <img src={companyLogo} style={{ width: '100%', height: '100%', objectFit: 'contain' }} alt="Preview" />
                        ) : (
                          <div style={{ fontSize: '20px', fontWeight: 'bold', color: 'var(--text-muted)' }}>POS</div>
                        )}
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <input 
                          type="file" 
                          id="company-logo-upload"
                          accept="image/*"
                          style={{ display: 'none' }}
                          onChange={(e) => {
                            const file = e.target.files[0];
                            if (file) {
                              const reader = new FileReader();
                              reader.onloadend = () => {
                                setCompanyLogo(reader.result);
                              };
                              reader.readAsDataURL(file);
                            }
                          }}
                        />
                        <button
                          type="button"
                          onClick={() => document.getElementById('company-logo-upload').click()}
                          style={{
                            padding: '8px 16px',
                            backgroundColor: 'var(--accent-color)',
                            color: '#fff',
                            border: 'none',
                            borderRadius: '8px',
                            fontWeight: '600',
                            fontSize: '12px',
                            cursor: 'pointer'
                          }}
                        >
                          {language === 'uz' ? "Logo yuklash" : "Загрузить лого"}
                        </button>
                        {companyLogo && (
                          <button
                            type="button"
                            onClick={() => setCompanyLogo(null)}
                            style={{
                              padding: '6px 12px',
                              backgroundColor: 'rgba(239, 68, 68, 0.1)',
                              color: 'var(--danger-color)',
                              border: 'none',
                              borderRadius: '8px',
                              fontWeight: '600',
                              fontSize: '11px',
                              cursor: 'pointer'
                            }}
                          >
                            {language === 'uz' ? "Yuklangan logoni o'chirish" : "Сбросить логотип"}
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Company Name */}
                  <div style={{ borderBottom: '1px solid var(--border-color)', paddingBottom: '20px' }}>
                    <label style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-primary)', display: 'block', marginBottom: '8px' }}>
                      {language === 'uz' ? "Firma Nomi" : "Название фирмы"}
                    </label>
                    <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                      <input 
                        type="text"
                        value={companyName}
                        onChange={(e) => setCompanyName(e.target.value)}
                        placeholder="Masalan: AGENT POS"
                        style={{ 
                          flexGrow: 1, 
                          padding: '10px 12px', 
                          borderRadius: '8px', 
                          border: '1px solid var(--border-color)', 
                          backgroundColor: 'var(--bg-primary)', 
                          color: 'var(--text-primary)', 
                          fontSize: '13px', 
                          fontWeight: '600' 
                        }}
                      />
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <input 
                          type="color"
                          value={companyNameColor}
                          onChange={(e) => setCompanyNameColor(e.target.value)}
                          style={{
                            width: '40px',
                            height: '38px',
                            padding: '2px',
                            borderRadius: '6px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            cursor: 'pointer'
                          }}
                        />
                        <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                          {language === 'uz' ? "Rang" : "Цвет"}
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Company Bio */}
                  <div style={{ borderBottom: '1px solid var(--border-color)', paddingBottom: '20px' }}>
                    <label style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-primary)', display: 'block', marginBottom: '8px' }}>
                      {language === 'uz' ? "Firma Sarlavhasi (Bio)" : "Подзаголовок фирмы (Био)"}
                    </label>
                    <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                      <input 
                        type="text"
                        value={companyBio}
                        onChange={(e) => setCompanyBio(e.target.value)}
                        placeholder="Masalan: MANAGEMENT"
                        style={{ 
                          flexGrow: 1, 
                          padding: '10px 12px', 
                          borderRadius: '8px', 
                          border: '1px solid var(--border-color)', 
                          backgroundColor: 'var(--bg-primary)', 
                          color: 'var(--text-primary)', 
                          fontSize: '13px', 
                          fontWeight: '600' 
                        }}
                      />
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <input 
                          type="color"
                          value={companyBioColor}
                          onChange={(e) => setCompanyBioColor(e.target.value)}
                          style={{
                            width: '40px',
                            height: '38px',
                            padding: '2px',
                            borderRadius: '6px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'transparent',
                            cursor: 'pointer'
                          }}
                        />
                        <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                          {language === 'uz' ? "Rang" : "Цвет"}
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Save Button */}
                  <button
                    onClick={() => {
                      localStorage.setItem('companyLogo', companyLogo || '');
                      localStorage.setItem('companyName', companyName);
                      localStorage.setItem('companyNameColor', companyNameColor);
                      localStorage.setItem('companyBio', companyBio);
                      localStorage.setItem('companyBioColor', companyBioColor);
                      showAlert(language === 'uz' ? "Firma sozlamalari saqlandi!" : "Настройки фирмы сохранены!", 'success');
                    }}
                    style={{
                      padding: '12px',
                      backgroundColor: 'var(--accent-color)',
                      color: '#fff',
                      border: 'none',
                      borderRadius: '8px',
                      fontWeight: '600',
                      fontSize: '14px',
                      cursor: 'pointer',
                      width: '100%',
                      marginTop: '8px'
                    }}
                  >
                    {language === 'uz' ? "Sozlamalarni saqlash" : "Сохранить настройки"}
                  </button>

                </div>
              </div>
            </div>
          )}


          {/* VIEW: ADMIN MANAGEMENT */}
          {activeTab === 'settings_admins' && (
            <div className="fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 2fr', gap: '24px' }}>
                
                {/* Add Admin Form */}
                <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', height: 'fit-content' }}>
                  <h3 style={{ fontSize: '16px', fontWeight: '500', marginBottom: '20px', color: 'var(--text-primary)' }}>
                    {language === 'uz' ? "Yangi Admin Qo'shish" : "Добавить нового админа"}
                  </h3>
                  <form onSubmit={handleAddAdmin} style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Login (Username)" : "Логин (Username)"}
                      </label>
                      <input 
                        type="text" 
                        required
                        value={newAdmin.login}
                        onChange={(e) => setNewAdmin({ ...newAdmin, login: e.target.value })}
                        placeholder={language === 'uz' ? 'Masalan: admin2' : 'Например: admin2'}
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600' }}
                      />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Parol" : "Пароль"}
                      </label>
                      <input 
                        type="password" 
                        required
                        value={newAdmin.password}
                        onChange={(e) => setNewAdmin({ ...newAdmin, password: e.target.value })}
                        placeholder="••••••••"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Ism (F.I.SH.)" : "Имя (Ф.И.О.)"}
                      </label>
                      <input 
                        type="text"
                        value={newAdmin.name}
                        onChange={(e) => setNewAdmin({ ...newAdmin, name: e.target.value })}
                        placeholder={language === 'uz' ? 'Masalan: Alisher Valiyev' : 'Например: Алишер Валиев'}
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Telefon" : "Телефон"}
                      </label>
                      <input 
                        type="text"
                        value={newAdmin.phone}
                        onChange={(e) => setNewAdmin({ ...newAdmin, phone: e.target.value })}
                        placeholder="+998 90 123 45 67"
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px' }}
                      />
                    </div>
                    <div>
                      <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                        {language === 'uz' ? "Rol" : "Роль"}
                      </label>
                      <select 
                        value={newAdmin.role}
                        onChange={(e) => setNewAdmin({ ...newAdmin, role: e.target.value })}
                        style={{ width: '100%', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', cursor: 'pointer' }}
                      >
                        <option value="admin">{language === 'uz' ? "Bosh Admin (Administrator)" : "Главный Админ (Администратор)"}</option>
                        <option value="warehouse_manager">{language === 'uz' ? "Ombor Mudiri" : "Заведующий складом"}</option>
                      </select>
                    </div>
                    <button 
                      type="submit"
                      style={{ width: '100%', padding: '12px', borderRadius: '8px', border: 'none', backgroundColor: 'var(--accent-color)', color: '#fff', fontWeight: '600', cursor: 'pointer', marginTop: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}
                    >
                      <Plus size={16} />
                      {language === 'uz' ? "Qo'shish" : "Добавить"}
                    </button>
                  </form>
                </div>

                {/* Admin List Panel */}
                <div style={{ backgroundColor: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--border-color)', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                  <h3 style={{ fontSize: '16px', fontWeight: '500', color: 'var(--text-primary)' }}>
                    {language === 'uz' ? "Administratorlar Ro'yxati" : "Список администраторов"}
                  </h3>
                  
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', overflowY: 'auto', maxHeight: '550px', paddingRight: '4px' }}>
                    {admins.map(adm => (
                      <div 
                        key={adm.id}
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          padding: '16px',
                          borderRadius: '8px',
                          backgroundColor: 'var(--bg-primary)',
                          border: '1px solid var(--border-color)',
                          transition: 'transform 0.2s, box-shadow 0.2s'
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                          {/* Avatar */}
                          <div style={{
                            width: '40px',
                            height: '40px',
                            borderRadius: '50%',
                            backgroundColor: adm.role === 'admin' ? 'rgba(13, 148, 136, 0.15)' : 'rgba(234, 179, 8, 0.15)',
                            color: adm.role === 'admin' ? 'var(--accent-color)' : 'var(--warning-color)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontWeight: '700',
                            fontSize: '14px'
                          }}>
                            {adm.role === 'admin' ? 'AD' : 'OM'}
                          </div>
                          <div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                              <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>
                                {adm.name || adm.username}
                              </span>
                              <span style={{
                                fontSize: '10px',
                                fontWeight: '700',
                                padding: '2px 8px',
                                borderRadius: '10px',
                                backgroundColor: adm.role === 'admin' ? 'rgba(13, 148, 136, 0.1)' : 'rgba(234, 179, 8, 0.1)',
                                color: adm.role === 'admin' ? 'var(--accent-color)' : 'var(--warning-color)'
                              }}>
                                {adm.role === 'admin' 
                                  ? (language === 'uz' ? 'Admin' : 'Админ') 
                                  : (language === 'uz' ? 'Ombor Mudiri' : 'Зав. складом')}
                              </span>
                              {adm.is_active === false && (
                                <span style={{
                                  fontSize: '10px',
                                  fontWeight: '700',
                                  padding: '2px 8px',
                                  borderRadius: '10px',
                                  backgroundColor: 'rgba(239, 68, 68, 0.1)',
                                  color: 'var(--danger-color)'
                                }}>
                                  {language === 'uz' ? 'Faol emas' : 'Неактивен'}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '2px' }}>
                              Username: <strong style={{ color: 'var(--text-primary)' }}>{adm.username}</strong> {adm.phone && ` | Tel: ${adm.phone}`}
                            </div>
                          </div>
                        </div>
                        
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <button 
                            onClick={() => {
                              setEditingAdmin({
                                ...adm,
                                password: ''
                              });
                              setShowEditAdminModal(true);
                            }}
                            style={{
                              padding: '8px',
                              borderRadius: '6px',
                              border: '1px solid var(--border-color)',
                              backgroundColor: 'transparent',
                              color: 'var(--text-secondary)',
                              cursor: 'pointer',
                              display: 'inline-flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              transition: 'all 0.2s'
                            }}
                            onMouseOver={(e) => e.currentTarget.style.color = 'var(--accent-color)'}
                            onMouseOut={(e) => e.currentTarget.style.color = 'var(--text-secondary)'}
                            title={language === 'uz' ? "Tahrirlash" : "Редактировать"}
                          >
                            <Edit size={14} />
                          </button>
                          <button 
                            onClick={() => handleDeleteAdmin(adm.id)}
                            disabled={adm.username === 'admin'}
                            style={{
                              padding: '8px',
                              borderRadius: '6px',
                              border: '1px solid var(--border-color)',
                              backgroundColor: 'transparent',
                              color: adm.username === 'admin' ? 'rgba(255,255,255,0.1)' : 'var(--danger-color)',
                              cursor: adm.username === 'admin' ? 'not-allowed' : 'pointer',
                              display: 'inline-flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              transition: 'all 0.2s'
                            }}
                            title={language === 'uz' ? "O'chirish" : "Удалить"}
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

              </div>
            </div>
          )}

        </div>
      </main>

      {/* Modal 1: Assign Tovar */}
      {showAssignProductModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999
        }} className="fade-in">
          <div style={{
            width: '400px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '16px',
            padding: '32px',
            boxShadow: 'var(--shadow-lg)'
          }}>
            <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '20px' }}>{t('assign_product_title')}</h3>
            <form onSubmit={handleAddAssignment} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <div>
                <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('select_product_label')}</label>
                <select 
                  required
                  value={newAssignment.productId}
                  onChange={(e) => setNewAssignment({ ...newAssignment, productId: e.target.value })}
                  style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', cursor: 'pointer' }}
                >
                  <option value="">{t('select_product_label')}...</option>
                  {products.map(prod => (
                    <option key={prod.id} value={prod.id}>{prod.name} ({prod.unit}) - Ombor: {prod.stock} {prod.unit}</option>
                  ))}
                </select>
              </div>

              <div>
                <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('qty_label')}</label>
                <input 
                  type="number" 
                  required
                  value={newAssignment.qty}
                  onChange={(e) => setNewAssignment({ ...newAssignment, qty: e.target.value })}
                  placeholder={t('qty_label')}
                  style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px' }}
                />
              </div>

              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                    {language === 'uz' ? "Amal qilish muddati (kun)" : "Срок действия (дней)"}
                  </label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', color: 'var(--text-secondary)', cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={newAssignment.isPermanent}
                      onChange={(e) => setNewAssignment({ 
                        ...newAssignment, 
                        isPermanent: e.target.checked,
                        durationDays: e.target.checked ? '9999' : '1'
                      })}
                      style={{ width: '14px', height: '14px', accentColor: 'var(--accent-color)', cursor: 'pointer' }}
                    />
                    {language === 'uz' ? "Doimiy" : "Постоянно"}
                  </label>
                </div>
                <input 
                  type="number"
                  min="1"
                  disabled={newAssignment.isPermanent}
                  value={newAssignment.isPermanent ? '' : newAssignment.durationDays}
                  onChange={(e) => setNewAssignment({ ...newAssignment, durationDays: e.target.value })}
                  placeholder={language === 'uz' ? "Kunlar soni..." : "Количество дней..."}
                  style={{ 
                    width: '100%', 
                    padding: '11px 14px', 
                    borderRadius: '8px', 
                    border: '1px solid var(--border-color)', 
                    backgroundColor: newAssignment.isPermanent ? 'var(--bg-tertiary)' : 'var(--bg-primary)', 
                    color: newAssignment.isPermanent ? 'var(--text-muted)' : 'var(--text-primary)', 
                    fontSize: '14px',
                    opacity: newAssignment.isPermanent ? 0.7 : 1
                  }}
                />
              </div>

              <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                <button 
                  type="button"
                  onClick={() => setShowAssignProductModal(false)}
                  style={{
                    flex: 1,
                    padding: '12px',
                    borderRadius: '8px',
                    border: '1px solid var(--border-color)',
                    backgroundColor: 'transparent',
                    color: 'var(--text-secondary)',
                    fontWeight: '600',
                    cursor: 'pointer'
                  }}
                >
                  {t('cancel')}
                </button>
                <button 
                  type="submit"
                  style={{
                    flex: 1,
                    padding: '12px',
                    borderRadius: '8px',
                    border: 'none',
                    backgroundColor: 'var(--accent-color)',
                    color: '#fff',
                    fontWeight: '600',
                    cursor: 'pointer'
                  }}
                >
                  {t('save')}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal 2: Assign Do'kon */}
      {showAssignStoreModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999
        }} className="fade-in">
          <div style={{
            width: '400px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '16px',
            padding: '32px',
            boxShadow: 'var(--shadow-lg)'
          }}>
            <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '20px' }}>{t('assign_store_title')}</h3>
            <form onSubmit={handleAddStoreAssignment} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <div>
                <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>{t('select_store_label')}</label>
                <select 
                  required
                  value={newStoreAssignment.storeId}
                  onChange={(e) => setNewStoreAssignment({ ...newStoreAssignment, storeId: e.target.value })}
                  style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', cursor: 'pointer' }}
                >
                  <option value="">{t('select_store_label')}...</option>
                  {stores.map(store => (
                    <option key={store.id} value={store.id}>{store.name}</option>
                  ))}
                </select>
              </div>

              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                  <label style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                    {language === 'uz' ? "Amal qilish muddati (kun)" : "Срок действия (дней)"}
                  </label>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', color: 'var(--text-secondary)', cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={newStoreAssignment.isPermanent}
                      onChange={(e) => setNewStoreAssignment({ 
                        ...newStoreAssignment, 
                        isPermanent: e.target.checked,
                        durationDays: e.target.checked ? '9999' : '1'
                      })}
                      style={{ width: '14px', height: '14px', accentColor: 'var(--accent-color)', cursor: 'pointer' }}
                    />
                    {language === 'uz' ? "Doimiy" : "Постоянно"}
                  </label>
                </div>
                <input 
                  type="number"
                  min="1"
                  disabled={newStoreAssignment.isPermanent}
                  value={newStoreAssignment.isPermanent ? '' : newStoreAssignment.durationDays}
                  onChange={(e) => setNewStoreAssignment({ ...newStoreAssignment, durationDays: e.target.value })}
                  placeholder={language === 'uz' ? "Kunlar soni..." : "Количество дней..."}
                  style={{ 
                    width: '100%', 
                    padding: '11px 14px', 
                    borderRadius: '8px', 
                    border: '1px solid var(--border-color)', 
                    backgroundColor: newStoreAssignment.isPermanent ? 'var(--bg-tertiary)' : 'var(--bg-primary)', 
                    color: newStoreAssignment.isPermanent ? 'var(--text-muted)' : 'var(--text-primary)', 
                    fontSize: '14px',
                    opacity: newStoreAssignment.isPermanent ? 0.7 : 1
                  }}
                />
              </div>

              <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                <button 
                  type="button"
                  onClick={() => setShowAssignStoreModal(false)}
                  style={{
                    flex: 1,
                    padding: '12px',
                    borderRadius: '8px',
                    border: '1px solid var(--border-color)',
                    backgroundColor: 'transparent',
                    color: 'var(--text-secondary)',
                    fontWeight: '600',
                    cursor: 'pointer'
                  }}
                >
                  {t('cancel')}
                </button>
                <button 
                  type="submit"
                  style={{
                    flex: 1,
                    padding: '12px',
                    borderRadius: '8px',
                    border: 'none',
                    backgroundColor: 'var(--accent-color)',
                    color: '#fff',
                    fontWeight: '600',
                    cursor: 'pointer'
                  }}
                >
                  {t('save')}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal: Edit Assigned Product Qty & Remaining Qty */}
       {showEditAssignmentModal && selectedAssignmentForEdit && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999
        }} className="fade-in">
          <div style={{
            width: '400px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '16px',
            padding: '32px',
            boxShadow: 'var(--shadow-lg)'
          }}>
            {userRole === 'agent' ? (
              // Read-only Details for Agent
              <div>
                <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '8px' }}>
                  {language === 'uz' ? 'Biriktirilgan mahsulot tafsilotlari' : 'Детали закрепленного товара'}
                </h3>
                <p style={{ fontSize: '14px', fontWeight: '600', color: 'var(--accent-color)', marginBottom: '24px' }}>
                  {selectedAssignmentForEdit.productName}
                </p>
                
                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginBottom: '24px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '12px', borderBottom: '1px solid var(--border-color)' }}>
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Olgan miqdor (Jami)' : 'Полученное количество (Всего)'}</span>
                    <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>{selectedAssignmentForEdit.qty} dona</span>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '12px', borderBottom: '1px solid var(--border-color)' }}>
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Qolgan miqdor (Qoldiq)' : 'Оставшееся количество (Остаток)'}</span>
                    <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--success-color)' }}>{selectedAssignmentForEdit.remainingQty !== undefined ? selectedAssignmentForEdit.remainingQty : selectedAssignmentForEdit.qty} dona</span>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{language === 'uz' ? 'Sana' : 'Дата'}</span>
                    <span style={{ fontSize: '14px', color: 'var(--text-primary)', fontWeight: '500' }}>{selectedAssignmentForEdit.date}</span>
                  </div>
                </div>

                <button 
                  onClick={() => {
                    setShowEditAssignmentModal(false);
                    setSelectedAssignmentForEdit(null);
                  }}
                  style={{
                    width: '100%',
                    padding: '12px',
                    borderRadius: '8px',
                    border: 'none',
                    backgroundColor: 'var(--accent-color)',
                    color: '#fff',
                    fontWeight: '600',
                    cursor: 'pointer',
                    transition: 'opacity 0.2s'
                  }}
                >
                  {language === 'uz' ? 'Yopish' : 'Закрыть'}
                </button>
              </div>
            ) : (
              // Original edit form for admin
              <>
                <h3 style={{ fontSize: '18px', fontWeight: '600', marginBottom: '8px' }}>
                  {language === 'uz' ? 'Biriktirilgan mahsulotni tahrirlash' : 'Редактировать закрепленный товар'}
                </h3>
                <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '20px' }}>
                  {selectedAssignmentForEdit.productName}
                </p>

                <form 
                  onSubmit={(e) => {
                    e.preventDefault();
                    const qtyVal = parseInt(editAssignmentQty);
                    const remVal = parseInt(editAssignmentRemainingQty);

                    if (isNaN(qtyVal) || qtyVal < 0) {
                      showAlert(language === 'uz' ? "Miqdor noto'g'ri kiritilgan" : "Некорректное количество", 'error');
                      return;
                    }
                    if (isNaN(remVal) || remVal < 0) {
                      showAlert(language === 'uz' ? "Qoldiq noto'g'ri kiritilgan" : "Некорректный остаток", 'error');
                      return;
                    }
                    if (remVal > qtyVal) {
                      showAlert(
                        language === 'uz' 
                          ? "Qolgan miqdor olgan miqdordan ko'p bo'lishi mumkin emas!" 
                          : "Остаток не может быть больше полученного количества!", 
                        'error'
                      );
                      return;
                    }

                    showConfirm(
                      language === 'uz'
                        ? "Ushbu mahsulotning biriktirilgan miqdorlarini o'zgartirishni tasdiqlaysizmi?"
                        : "Вы подтверждаете изменение закрепленных количеств этого товара?",
                      () => {
                        const success = handleUpdateAssignment(selectedAssignmentForEdit.id, qtyVal, remVal);
                        if (success) {
                          setShowEditAssignmentModal(false);
                          setSelectedAssignmentForEdit(null);
                        }
                      }
                    );
                  }} 
                  style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}
                >
                  <div>
                    <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                      {language === 'uz' ? 'Olgan miqdor (Jami)' : 'Полученное количество (Всего)'}
                    </label>
                    <input 
                      type="number" 
                      min="0"
                      required
                      value={editAssignmentQty}
                      onChange={(e) => setEditAssignmentQty(e.target.value)}
                      style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', fontWeight: '600' }}
                    />
                  </div>

                  <div>
                    <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                      {language === 'uz' ? 'Qolgan miqdor (Qoldiq)' : 'Оставшееся количество (Остаток)'}
                    </label>
                    <input 
                      type="number" 
                      min="0"
                      max={editAssignmentQty}
                      required
                      value={editAssignmentRemainingQty}
                      onChange={(e) => setEditAssignmentRemainingQty(e.target.value)}
                      style={{ width: '100%', padding: '11px 14px', borderRadius: '8px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '14px', fontWeight: '600' }}
                    />
                  </div>

                  <div style={{ display: 'flex', gap: '12px', marginTop: '12px' }}>
                    <button 
                      type="button"
                      onClick={() => {
                        setShowEditAssignmentModal(false);
                        setSelectedAssignmentForEdit(null);
                      }}
                      style={{
                        flex: 1,
                        padding: '12px',
                        borderRadius: '8px',
                        border: '1px solid var(--border-color)',
                        backgroundColor: 'transparent',
                        color: 'var(--text-secondary)',
                        fontWeight: '600',
                        cursor: 'pointer'
                      }}
                    >
                      {t('cancel')}
                    </button>
                    <button 
                      type="submit"
                      style={{
                        flex: 1,
                        padding: '12px',
                        borderRadius: '8px',
                        border: 'none',
                        backgroundColor: 'var(--accent-color)',
                        color: '#fff',
                        fontWeight: '600',
                        cursor: 'pointer'
                      }}
                    >
                      {language === 'uz' ? 'Tasdiqlash' : 'Подтвердить'}
                    </button>
                  </div>
                </form>
              </>
            )}
          </div>
        </div>
      )}
      {/* Modal: Product Agent Stock Details */}
      {selectedProductStockDetails && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999
        }} className="fade-in">
          <div style={{
            width: '550px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '16px',
            padding: '32px',
            boxShadow: 'var(--shadow-lg)',
            display: 'flex',
            flexDirection: 'column',
            gap: '20px'
          }}>
            <div>
              <h3 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)' }}>{selectedProductStockDetails}</h3>
              <p style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>{t('agent_stock_details_title')}</p>
            </div>

            <div style={{ maxHeight: '300px', overflowY: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                <thead>
                  <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '12px' }}>
                    <th style={{ padding: '8px 4px' }}>{t('agent_login_col')}</th>
                    <th style={{ padding: '8px 4px' }}>Username</th>
                    <th style={{ padding: '8px 4px', textAlign: 'right' }}>{t('given_col')}</th>
                    <th style={{ padding: '8px 4px', textAlign: 'right' }}>{t('remaining_col')}</th>
                  </tr>
                </thead>
                <tbody>
                  {assignments
                    .filter(ass => ass.productName === selectedProductStockDetails)
                    .map(ass => {
                      const agentObj = agents.find(a => a.id === ass.agentId);
                      return (
                        <tr key={ass.id} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '13px' }}>
                          <td style={{ padding: '10px 4px', fontWeight: '600', color: 'var(--text-primary)' }}>{getAgentLoginByName(ass.agentName)}</td>
                          <td style={{ padding: '10px 4px', color: 'var(--text-secondary)' }}>
                            {agentObj && agentObj.username ? `@${agentObj.username.replace(/^@/, '')}` : t('no_username')}
                          </td>
                          <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600' }}>{ass.qty} {language === 'uz' ? 'dona' : 'шт'}</td>
                          <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '700', color: (ass.remainingQty === "" ? 0 : (ass.remainingQty !== undefined ? ass.remainingQty : ass.qty)) > 0 ? 'var(--success-color)' : 'var(--text-muted)' }}>
                            {ass.remainingQty === "" ? 0 : (ass.remainingQty !== undefined ? ass.remainingQty : ass.qty)} {language === 'uz' ? 'dona' : 'шт'}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '8px' }}>
              <button 
                type="button"
                onClick={() => setSelectedProductStockDetails(null)}
                style={{
                  padding: '10px 24px',
                  borderRadius: '8px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: 'transparent',
                  color: 'var(--text-secondary)',
                  fontWeight: '600',
                  cursor: 'pointer'
                }}
              >
                {t('close')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Agent Store Sales Details */}
      {selectedTahlilAgentStore && (() => {
        const selAgent = agents.find(a => a.id === selectedTahlilAgentId);
        if (!selAgent) return null;

        const agentStoreSales = getFilteredSales(sales).filter(
          s => (s.agentId === selAgent.id || s.agent === selAgent.name) && s.store === selectedTahlilAgentStore
        );

        // Group by product name
        const productDetails = {};
        let totalRevenue = 0;
        let totalProfit = 0;

        agentStoreSales.forEach(s => {
          if (s.items) {
            s.items.forEach(item => {
              const cost = item.originalPrice || (products.find(p => p.name === item.productName)?.originalPrice) || 0;
              const revenue = item.qty * item.price;
              const profit = item.qty * (item.price - cost);

              if (!productDetails[item.productName]) {
                productDetails[item.productName] = { qty: 0, revenue: 0, profit: 0 };
              }
              productDetails[item.productName].qty += item.qty;
              productDetails[item.productName].revenue += revenue;
              productDetails[item.productName].profit += profit;

              totalRevenue += revenue;
              totalProfit += profit;
            });
          }
        });

        return (
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100vw',
            height: '100vh',
            backgroundColor: 'rgba(15, 23, 42, 0.85)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 9999
          }} className="fade-in">
            <div style={{
              width: '650px',
              backgroundColor: 'var(--bg-secondary)',
              border: '1px solid var(--border-color)',
              borderRadius: '16px',
              padding: '32px',
              boxShadow: 'var(--shadow-lg)',
              display: 'flex',
              flexDirection: 'column',
              gap: '20px'
            }}>
              <div>
                <h3 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)' }}>
                  {t('store_sales_details_title').replace('{store}', selectedTahlilAgentStore)}
                </h3>
                <p style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>
                  {t('store_sales_details_desc').replace('{agent}', selAgent.login).replace('{username}', selAgent.username || 'username')}
                </p>
              </div>

              <div style={{ maxHeight: '300px', overflowY: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '12px' }}>
                      <th style={{ padding: '8px 4px' }}>{t('product_name')}</th>
                      <th style={{ padding: '8px 4px', textAlign: 'right' }}>{language === 'uz' ? 'Soni' : 'Кол-во'}</th>
                      <th style={{ padding: '8px 4px', textAlign: 'right' }}>{t('revenue')}</th>
                      <th style={{ padding: '8px 4px', textAlign: 'right' }}>{t('net_profit')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Object.keys(productDetails).length === 0 ? (
                      <tr>
                        <td colSpan="4" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                          {language === 'uz' ? 'Mahsulot sotilmagan.' : 'Товар не продан.'}
                        </td>
                      </tr>
                    ) : (
                      Object.keys(productDetails).map(prodName => {
                        const info = productDetails[prodName];
                        return (
                          <tr key={prodName} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '13px' }}>
                            <td style={{ padding: '10px 4px', fontWeight: '500', color: 'var(--text-primary)' }}>{prodName}</td>
                            <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600' }}>{info.qty.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</td>
                            <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600', color: 'var(--accent-color)' }}>{info.revenue.toLocaleString()} UZS</td>
                            <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600', color: 'var(--success-color)' }}>{info.profit.toLocaleString()} UZS</td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {/* Total Summary Row in modal */}
              <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                backgroundColor: 'var(--bg-primary)',
                padding: '16px',
                borderRadius: '8px',
                border: '1px solid var(--border-color)',
                marginTop: '8px'
              }}>
                <div>
                  <div style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>{t('tushum_badge')}</div>
                  <div style={{ fontSize: '15px', fontWeight: '700', color: 'var(--accent-color)' }}>{totalRevenue.toLocaleString()} UZS</div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <div style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>{t('profit_badge')}</div>
                  <div style={{ fontSize: '15px', fontWeight: '700', color: 'var(--success-color)' }}>{totalProfit.toLocaleString()} UZS</div>
                </div>
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '8px' }}>
                <button 
                  type="button"
                  onClick={() => setSelectedTahlilAgentStore(null)}
                  style={{
                    padding: '10px 24px',
                    borderRadius: '8px',
                    border: '1px solid var(--border-color)',
                    backgroundColor: 'transparent',
                    color: 'var(--text-secondary)',
                    fontWeight: '600',
                    cursor: 'pointer'
                  }}
                >
                  {t('close')}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Modal: Product Distribution Details */}
      {selectedTahlilProductDetails && (() => {
        const filteredSales = getFilteredSales(sales);
        const distributionList = [];
        let totalSoldQty = 0;
        let totalRevenue = 0;
        const uniqueStoresSet = new Set();

        filteredSales.forEach(s => {
          if (s.items) {
            s.items.forEach(item => {
              if (item.productName === selectedTahlilProductDetails) {
                // Find agent details
                const agentObj = agents.find(a => a.name === s.agent || a.login === s.agent);
                distributionList.push({
                  storeName: s.store,
                  qty: item.qty,
                  price: item.price,
                  revenue: item.qty * item.price,
                  agentLogin: agentObj ? agentObj.login : s.agent,
                  agentUsername: agentObj ? agentObj.username : 'username',
                  date: s.date
                });
                uniqueStoresSet.add(s.store);
                totalSoldQty += item.qty;
                totalRevenue += (item.qty * item.price);
              }
            });
          }
        });

        return (
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100vw',
            height: '100vh',
            backgroundColor: 'rgba(15, 23, 42, 0.85)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 9999
          }} className="fade-in">
            <div style={{
              width: '700px',
              backgroundColor: 'var(--bg-secondary)',
              border: '1px solid var(--border-color)',
              borderRadius: '16px',
              padding: '32px',
              boxShadow: 'var(--shadow-lg)',
              display: 'flex',
              flexDirection: 'column',
              gap: '20px'
            }}>
              <div>
                <h3 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)' }}>
                  {t('product_dist_title').replace('{product}', selectedTahlilProductDetails)}
                </h3>
                <p style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>
                  {t('product_dist_desc')}
                </p>
              </div>

              {/* Summary Badges */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                <div style={{ backgroundColor: 'var(--bg-primary)', padding: '12px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <span style={{ fontSize: '11px', color: 'var(--text-secondary)', display: 'block', marginBottom: '4px' }}>{t('total_sold_badge')}</span>
                  <span style={{ fontSize: '16px', fontWeight: '700', color: 'var(--accent-color)' }}>{totalSoldQty.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</span>
                </div>
                <div style={{ backgroundColor: 'var(--bg-primary)', padding: '12px 16px', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <span style={{ fontSize: '11px', color: 'var(--text-secondary)', display: 'block', marginBottom: '4px' }}>{t('total_stores_badge')}</span>
                  <span style={{ fontSize: '16px', fontWeight: '700', color: 'var(--success-color)' }}>{uniqueStoresSet.size} {language === 'uz' ? 'ta do\'kon' : 'магазинов'}</span>
                </div>
              </div>

              <div style={{ maxHeight: '250px', overflowY: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ borderBottom: '1px solid var(--border-color)', color: 'var(--text-secondary)', fontSize: '12px' }}>
                      <th style={{ padding: '8px 4px' }}>{language === 'uz' ? 'Do\'kon nomi' : 'Название магазина'}</th>
                      <th style={{ padding: '8px 4px', textAlign: 'right' }}>{language === 'uz' ? 'Soni' : 'Кол-во'}</th>
                      <th style={{ padding: '8px 4px', textAlign: 'right' }}>{t('revenue')}</th>
                      <th style={{ padding: '8px 4px' }}>{language === 'uz' ? 'Agent' : 'Агент'}</th>
                      <th style={{ padding: '8px 4px', textAlign: 'right' }}>{t('date')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {distributionList.length === 0 ? (
                      <tr>
                        <td colSpan="5" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                          {t('no_dist_found')}
                        </td>
                      </tr>
                    ) : (
                      distributionList.map((dist, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--border-color)', fontSize: '13px' }}>
                          <td style={{ padding: '10px 4px', fontWeight: '500', color: 'var(--text-primary)' }}>{dist.storeName}</td>
                          <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600' }}>{dist.qty.toLocaleString()} {language === 'uz' ? 'dona' : 'шт'}</td>
                          <td style={{ padding: '10px 4px', textAlign: 'right', fontWeight: '600', color: 'var(--accent-color)' }}>{dist.revenue.toLocaleString()} UZS</td>
                          <td style={{ padding: '10px 4px' }}>
                            <div style={{ fontWeight: '600', color: 'var(--text-primary)' }}>{dist.agentLogin}</div>
                            <div style={{ fontSize: '10px', color: 'var(--text-secondary)' }}>@{dist.agentUsername}</div>
                          </td>
                          <td style={{ padding: '10px 4px', textAlign: 'right', color: 'var(--text-secondary)' }}>{dist.date}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '8px' }}>
                <button 
                  type="button"
                  onClick={() => setSelectedTahlilProductDetails(null)}
                  style={{
                    padding: '10px 24px',
                    borderRadius: '8px',
                    border: '1px solid var(--border-color)',
                    backgroundColor: 'transparent',
                    color: 'var(--text-secondary)',
                    fontWeight: '600',
                    cursor: 'pointer'
                  }}
                >
                  {t('close')}
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Custom React Toast Notification */}
            {/* Kassa (POS) Terminal Overlay */}
            {activeCashierStore && userRole === 'agent' && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'var(--bg-primary)',
          color: 'var(--text-primary)',
          zIndex: 10000,
          display: 'flex',
          flexDirection: 'column',
          boxSizing: 'border-box',
          overflow: 'hidden'
        }} className="fade-in">
          {/* Header */}
          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            padding: '16px 20px',
            borderBottom: '1px solid var(--border-color)',
            backgroundColor: 'var(--bg-secondary)',
            flexShrink: 0
          }}>
            <div>
              <h2 style={{ fontSize: '16px', fontWeight: '700', color: 'var(--text-primary)' }}>
                {activeCashierStore.storeName}
              </h2>
              <p style={{ fontSize: '11px', color: 'var(--text-secondary)', margin: 0 }}>
                👤 {activeCashierStore.ownerName || 'Tadbirkor'} • 📍 {activeCashierStore.address || 'Manzil kiritilmagan'}
              </p>
            </div>
            <button
              onClick={handleExitCashier}
              style={{
                border: '1px solid var(--border-color)',
                backgroundColor: 'rgba(239, 68, 68, 0.1)',
                color: '#ef4444',
                padding: '6px 12px',
                borderRadius: '6px',
                cursor: 'pointer',
                fontWeight: '600',
                fontSize: '12px'
              }}
            >
              {language === 'uz' ? "Chiqish" : "Выход"}
            </button>
          </div>

          {!showPaymentSection ? (
            /* ==================== SCREEN 1: CART & PRODUCT ADDITION ==================== */
            <div style={{
              display: 'flex',
              flexDirection: 'column',
              flexGrow: 1,
              overflow: 'hidden',
              padding: '16px',
              gap: '16px'
            }}>
              {/* Barcode & Search Trigger Inputs */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', flexShrink: 0 }}>
                <div className="cashier-input-row flex-nowrap-mobile" style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                  {/* Barcode Input Form */}
                  <form onSubmit={handleBarcodeSubmit} className="flex-nowrap-mobile" style={{ display: 'flex', gap: '8px', flexGrow: 1 }}>
                    <input
                      type="text"
                      placeholder={language === 'uz' ? "Shtrix-kod..." : "Штрих-код..."}
                      value={barcodeInput}
                      onChange={(e) => setBarcodeInput(e.target.value)}
                      autoFocus
                      style={{
                        flexGrow: 1,
                        padding: '12px',
                        borderRadius: '8px',
                        border: '1px solid var(--border-color)',
                        backgroundColor: 'var(--bg-secondary)',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        boxSizing: 'border-box',
                        minWidth: 0
                      }}
                    />
                    <button
                      type="submit"
                      style={{
                        padding: '12px 16px',
                        backgroundColor: 'var(--accent-color)',
                        color: '#fff',
                        border: 'none',
                        borderRadius: '8px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        fontSize: '13px',
                        whiteSpace: 'nowrap'
                      }}
                    >
                      {language === 'uz' ? "Qo'shish" : "Добавить"}
                    </button>
                  </form>

                  {/* Camera Scanner Trigger */}
                  <button
                    type="button"
                    onClick={() => setShowCameraScanner(true)}
                    style={{
                      width: '42px',
                      height: '42px',
                      borderRadius: '8px',
                      border: '1px solid var(--border-color)',
                      backgroundColor: 'var(--bg-secondary)',
                      color: 'var(--text-primary)',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      cursor: 'pointer',
                      flexShrink: 0
                    }}
                    title={language === 'uz' ? "Kamera orqali skanerlash" : "Сканировать камерой"}
                  >
                    <Camera size={20} />
                  </button>

                  {/* Product Selector Trigger (+) */}
                  <button
                    type="button"
                    onClick={() => setShowProductModal(true)}
                    style={{
                      width: '42px',
                      height: '42px',
                      borderRadius: '8px',
                      border: 'none',
                      backgroundColor: 'var(--accent-color)',
                      color: '#fff',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      cursor: 'pointer',
                      flexShrink: 0,
                      fontSize: '22px',
                      fontWeight: 'bold',
                      lineHeight: '42px'
                    }}
                    title={language === 'uz' ? "Mahsulot qo'shish" : "Добавить товар"}
                  >
                    +
                  </button>
                </div>
              </div>

              {/* Shopping Cart List */}
              <div style={{
                flexGrow: 1,
                overflowY: 'auto',
                border: '1px solid var(--border-color)',
                borderRadius: '8px',
                backgroundColor: 'var(--bg-secondary)',
                padding: '12px',
                display: 'flex',
                flexDirection: 'column',
                gap: '8px'
              }}>
                <h3 style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-secondary)', margin: '0 0 8px 0', borderBottom: '1px solid var(--border-color)', paddingBottom: '6px' }}>
                  {language === 'uz' ? "Sotuv savatchasi" : "Корзина продаж"}
                </h3>
                {cashierCart.length === 0 ? (
                  <div style={{ margin: 'auto', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13px' }}>
                    {language === 'uz' ? "Savatcha bo'sh. Mahsulot qo'shing." : "Корзина пуста. Добавьте товары."}
                  </div>
                ) : (
                  cashierCart.map(item => (
                    <div
                      key={item.productId}
                      className="cart-item-row flex-nowrap-mobile"
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        padding: '10px 0',
                        borderBottom: '1px solid var(--border-color)'
                      }}
                    >
                      <div className="cart-item-info" style={{ display: 'flex', flexDirection: 'column', gap: '2px', maxWidth: '50%' }}>
                        <span style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-primary)', wordBreak: 'break-word' }}>
                          {item.productName}
                        </span>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                          <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                            {parseFloat(item.price).toLocaleString('uz-UZ')} so'm
                          </span>
                          <span style={{ 
                            fontSize: '10px', 
                            fontWeight: '600', 
                            color: 'var(--success-color)', 
                            backgroundColor: 'var(--success-light)',
                            padding: '1px 6px',
                            borderRadius: '4px'
                          }}>
                            {language === 'uz' ? `Qoldiq: ${item.maxQty - item.quantity} ${item.unit || 'dona'}` : `Ост: ${item.maxQty - item.quantity} ${item.unit || 'шт'}`}
                          </span>
                        </div>
                      </div>

                      {/* Quantity Editor */}
                      <div className="cart-item-qty-container flex-nowrap-mobile" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <button
                          className="cart-item-qty-btn"
                          onClick={() => {
                            if (item.quantity > 1) {
                              handleUpdateCartItemQty(item.productId, item.quantity - 1);
                            }
                          }}
                          style={{
                            width: '28px',
                            height: '28px',
                            borderRadius: '4px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'var(--bg-primary)',
                            color: 'var(--text-primary)',
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontWeight: 'bold'
                          }}
                        >
                          -
                        </button>
                        <input
                          type="number"
                          className="cart-item-qty-input"
                          value={item.quantity}
                          onChange={(e) => handleUpdateCartItemQty(item.productId, e.target.value)}
                          style={{
                            width: '45px',
                            textAlign: 'center',
                            padding: '4px',
                            border: '1px solid var(--border-color)',
                            borderRadius: '4px',
                            backgroundColor: 'var(--bg-primary)',
                            color: 'var(--text-primary)',
                            fontSize: '13px',
                            fontWeight: '600'
                          }}
                        />
                        <button
                          className="cart-item-qty-btn"
                          onClick={() => {
                            handleUpdateCartItemQty(item.productId, item.quantity + 1);
                          }}
                          style={{
                            width: '28px',
                            height: '28px',
                            borderRadius: '4px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: 'var(--bg-primary)',
                            color: 'var(--text-primary)',
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontWeight: 'bold'
                          }}
                        >
                          +
                        </button>
                      </div>

                      {/* Subtotal & Delete */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <span className="cart-item-subtotal" style={{ fontSize: '13px', fontWeight: '700', color: 'var(--text-primary)', textAlign: 'right', minWidth: '80px' }}>
                          {(item.price * item.quantity).toLocaleString('uz-UZ')} so'm
                        </span>
                        <button
                          onClick={() => handleRemoveCartItem(item.productId)}
                          style={{
                            border: 'none',
                            backgroundColor: 'transparent',
                            color: '#ef4444',
                            cursor: 'pointer',
                            padding: '4px'
                          }}
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Total display & To'lov button */}
              <div style={{
                padding: '16px',
                border: '1px solid var(--border-color)',
                borderRadius: '8px',
                backgroundColor: 'var(--bg-secondary)',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                flexShrink: 0
              }}>
                <div>
                  <span style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block' }}>
                    {language === 'uz' ? "Jami summa" : "Итоговая сумма"}
                  </span>
                  <span style={{ fontSize: '20px', fontWeight: '800', color: 'var(--accent-color)' }}>
                    {cashierCart.reduce((acc, item) => acc + (item.price * item.quantity), 0).toLocaleString('uz-UZ')} so'm
                  </span>
                </div>
                <button
                  onClick={() => {
                    if (cashierCart.length === 0) {
                      showAlert(language === 'uz' ? "Savatcha bo'sh!" : "Корзина пуста!", 'error');
                      return;
                    }
                    setShowPaymentSection(true);
                  }}
                  style={{
                    padding: '14px 28px',
                    backgroundColor: 'var(--accent-color)',
                    color: '#fff',
                    border: 'none',
                    borderRadius: '8px',
                    fontWeight: '700',
                    cursor: 'pointer',
                    fontSize: '15px',
                    transition: 'all 0.2s ease',
                    boxShadow: '0 4px 10px rgba(13, 148, 136, 0.3)'
                  }}
                >
                  {language === 'uz' ? "To'lovga o'tish" : "Перейти к оплате"} →
                </button>
              </div>
            </div>
          ) : (
            /* ==================== SCREEN 2: PAYMENT & DISCOUNT SELECTION ==================== */
            <div style={{
              display: 'flex',
              flexDirection: 'column',
              flexGrow: 1,
              overflowY: 'auto',
              padding: '16px',
              gap: '20px'
            }}>
              {/* Discount Section */}
              <div style={{
                backgroundColor: 'var(--bg-secondary)',
                padding: '16px',
                borderRadius: '8px',
                border: '1px solid var(--border-color)'
              }}>
                <h3 style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)', marginBottom: '12px' }}>
                  {language === 'uz' ? "Chegirma tanlash" : "Выбор скидки"}
                </h3>
                
                {/* Predefined discount buttons */}
                <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', marginBottom: '14px' }}>
                  {[0, ...discountsList].map(disc => {
                    const isSelected = !customDiscountInput && cashierDiscount === disc;
                    return (
                      <button
                        key={disc}
                        onClick={() => {
                          setCashierDiscount(disc);
                          setCustomDiscountInput('');
                        }}
                        style={{
                          padding: '10px 16px',
                          borderRadius: '6px',
                          border: '1px solid var(--border-color)',
                          backgroundColor: isSelected ? 'var(--accent-color)' : 'var(--bg-primary)',
                          color: isSelected ? '#fff' : 'var(--text-primary)',
                          fontWeight: '600',
                          fontSize: '13px',
                          cursor: 'pointer',
                          transition: 'all 0.2s ease'
                        }}
                      >
                        {disc === 0 ? (language === 'uz' ? "Chegirmasiz" : "Без скидки") : `${disc}%`}
                      </button>
                    );
                  })}
                </div>

                {/* Custom discount input */}
                {customDiscountEnabled && (
                  <div>
                    <label style={{ fontSize: '12px', color: 'var(--text-secondary)', display: 'block', marginBottom: '6px' }}>
                      {language === 'uz' ? "Ixtiyoriy chegirma foizi (%)" : "Произвольный процент скидки (%)"}
                    </label>
                    <input
                      type="number"
                      min="0"
                      max="99"
                      placeholder="Masalan: 12"
                      value={customDiscountInput}
                      onChange={(e) => {
                        setCustomDiscountInput(e.target.value);
                        setCashierDiscount(0);
                      }}
                      style={{
                        width: '100%',
                        padding: '10px',
                        borderRadius: '6px',
                        border: '1px solid var(--border-color)',
                        backgroundColor: 'var(--bg-primary)',
                        color: 'var(--text-primary)',
                        fontSize: '13px',
                        boxSizing: 'border-box'
                      }}
                    />
                  </div>
                )}
              </div>

              {/* Payment Method Section */}
              <div style={{
                backgroundColor: 'var(--bg-secondary)',
                padding: '16px',
                borderRadius: '8px',
                border: '1px solid var(--border-color)'
              }}>
                <h3 style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)', marginBottom: '12px' }}>
                  {language === 'uz' ? "To'lov turi" : "Способ оплаты"}
                </h3>
                <div style={{ 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px', 
                  padding: '14px 16px', 
                  borderRadius: '8px', 
                  border: '2px solid var(--accent-color)', 
                  backgroundColor: 'rgba(13, 148, 136, 0.08)' 
                }}>
                  <CreditCard size={20} style={{ color: 'var(--accent-color)' }} />
                  <div>
                    <span style={{ fontWeight: '700', fontSize: '14px', color: 'var(--text-primary)', display: 'block' }}>
                      Tinda Terminal Payment
                    </span>
                    <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                      {language === 'uz' ? "To'lov avtomatik ravishda smart-terminal orqali qabul qilinadi." : "Оплата автоматически принимается через смарт-терминал."}
                    </span>
                  </div>
                </div>
              </div>

              {/* Order Summary & Confirm */}
              {(() => {
                const subtotal = cashierCart.reduce((acc, item) => acc + (item.price * item.quantity), 0);
                const discountPercentage = customDiscountEnabled && customDiscountInput ? parseFloat(customDiscountInput) || 0 : cashierDiscount;
                const discountAmount = Math.round(subtotal * (discountPercentage / 100));
                const finalTotal = subtotal - discountAmount;

                return (
                  <div style={{
                    marginTop: 'auto',
                    backgroundColor: 'var(--bg-secondary)',
                    padding: '20px',
                    borderRadius: '8px',
                    border: '1px solid var(--border-color)',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '12px'
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: 'var(--text-secondary)' }}>
                      <span>{language === 'uz' ? "Savatcha summasi" : "Сумма корзины"}</span>
                      <span>{subtotal.toLocaleString('uz-UZ')} so'm</span>
                    </div>
                    {discountAmount > 0 && (
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: '#ef4444' }}>
                        <span>{language === 'uz' ? `Chegirma (${discountPercentage}%)` : `Скидка (${discountPercentage}%)`}</span>
                        <span>-{discountAmount.toLocaleString('uz-UZ')} so'm</span>
                      </div>
                    )}
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '16px', fontWeight: '800', color: 'var(--text-primary)', borderTop: '1px solid var(--border-color)', paddingTop: '10px' }}>
                      <span>{language === 'uz' ? "To'lanishi kerak" : "К оплате"}</span>
                      <span style={{ color: 'var(--accent-color)' }}>{finalTotal.toLocaleString('uz-UZ')} so'm</span>
                    </div>

                    <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
                      <button
                        onClick={() => setShowPaymentSection(false)}
                        style={{
                          flexGrow: 1,
                          padding: '14px',
                          border: '1px solid var(--border-color)',
                          backgroundColor: 'var(--bg-primary)',
                          color: 'var(--text-primary)',
                          borderRadius: '8px',
                          fontWeight: '600',
                          cursor: 'pointer',
                          fontSize: '14px'
                        }}
                      >
                        ← {language === 'uz' ? "Ortga" : "Назад"}
                      </button>
                      <button
                        onClick={() => {
                          handleTindaPayment(subtotal, discountAmount, finalTotal);
                        }}
                        disabled={!tindaTerminalIp}
                        style={{
                          flexGrow: 2,
                          padding: '14px',
                          backgroundColor: !tindaTerminalIp ? 'var(--text-muted)' : 'var(--accent-color)',
                          color: '#fff',
                          border: 'none',
                          borderRadius: '8px',
                          fontWeight: '700',
                          cursor: !tindaTerminalIp ? 'not-allowed' : 'pointer',
                          fontSize: '14px',
                          boxShadow: '0 4px 10px rgba(13, 148, 136, 0.3)'
                        }}
                      >
                        ✓ {language === 'uz' ? "Tinda orqali to'lash" : "Оплатить через Tinda"}
                      </button>
                    </div>
                  </div>
                );
              })()}
            </div>
          )}

          {/* CAMERA SCANNER MODAL */}
          {showCameraScanner && (
            <div style={{
              position: 'fixed',
              top: 0,
              left: 0,
              width: '100vw',
              height: '100vh',
              backgroundColor: 'rgba(15, 23, 42, 0.9)',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 11000,
              padding: '20px',
              boxSizing: 'border-box'
            }} className="fade-in">
              <div 
                id="camera-reader-container"
                style={{
                  width: '100%',
                  maxWidth: '400px',
                  backgroundColor: 'var(--bg-secondary)',
                  borderRadius: '12px',
                  border: '1px solid var(--border-color)',
                  padding: '16px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '16px',
                  position: 'relative',
                  boxShadow: '0 10px 25px rgba(0,0,0,0.3)',
                  transition: 'box-shadow 0.3s ease'
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <h3 style={{ fontSize: '15px', fontWeight: '700', margin: 0, color: 'var(--text-primary)' }}>
                    {language === 'uz' ? "Kamera orqali skanerlash" : "Сканирование камерой"}
                  </h3>
                  <button 
                    onClick={() => setShowCameraScanner(false)}
                    style={{
                      border: 'none',
                      backgroundColor: 'transparent',
                      color: 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '20px',
                      lineHeight: 1
                    }}
                  >
                    &times;
                  </button>
                </div>

                {/* Scanner Frame */}
                <div style={{ 
                  position: 'relative', 
                  width: '100%', 
                  paddingTop: '100%', 
                  overflow: 'hidden', 
                  borderRadius: '8px',
                  backgroundColor: '#000'
                }}>
                  <div 
                    id="camera-reader" 
                    style={{ 
                      position: 'absolute', 
                      top: 0, 
                      left: 0, 
                      width: '100%', 
                      height: '100%' 
                    }} 
                  />
                  <div style={{
                    position: 'absolute',
                    top: 0,
                    left: 0,
                    width: '100%',
                    height: '3px',
                    backgroundColor: 'var(--accent-color)',
                    boxShadow: '0 0 10px var(--accent-color)',
                    animation: 'scanAnimation 2s linear infinite',
                    zIndex: 10
                  }} />
                </div>

                <p style={{ fontSize: '11px', color: 'var(--text-secondary)', textAlign: 'center', margin: 0 }}>
                  {language === 'uz' ? "Mahsulot shtrix-kodini kamera markaziga qarating." : "Направьте штрих-код товара в центр камеры."}
                </p>

                <button
                  onClick={() => setShowCameraScanner(false)}
                  style={{
                    width: '100%',
                    padding: '12px',
                    backgroundColor: 'rgba(239, 68, 68, 0.1)',
                    color: '#ef4444',
                    border: '1px solid rgba(239, 68, 68, 0.2)',
                    borderRadius: '8px',
                    fontWeight: '600',
                    cursor: 'pointer',
                    fontSize: '13px'
                  }}
                >
                  {language === 'uz' ? "Yopish" : "Закрыть"}
                </button>
              </div>
              <style dangerouslySetInnerHTML={{__html: `
                @keyframes scanAnimation {
                  0% { top: 0%; }
                  50% { top: 100%; }
                  100% { top: 0%; }
                }
              `}} />
            </div>
          )}


          {/* PRODUCT SELECTION MODAL */}
          {showProductModal && (
            <div style={{
              position: 'fixed',
              top: 0,
              left: 0,
              width: '100vw',
              height: '100vh',
              backgroundColor: 'rgba(15, 23, 42, 0.85)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              zIndex: 11000,
              padding: '16px',
              boxSizing: 'border-box'
            }} className="fade-in">
              <div style={{
                width: '100%',
                maxWidth: '450px',
                maxHeight: '85vh',
                backgroundColor: 'var(--bg-secondary)',
                borderRadius: '12px',
                border: '1px solid var(--border-color)',
                display: 'flex',
                flexDirection: 'column',
                overflow: 'hidden',
                boxShadow: 'var(--shadow-xl)'
              }}>
                <div style={{
                  padding: '16px 20px',
                  borderBottom: '1px solid var(--border-color)',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  flexShrink: 0
                }}>
                  <h3 style={{ fontSize: '15px', fontWeight: '700', color: 'var(--text-primary)', margin: 0 }}>
                    {language === 'uz' ? "Mahsulotlarni tanlash" : "Выбор товаров"}
                  </h3>
                  <button
                    onClick={() => {
                      setSelectedModalProducts([]);
                      setShowProductModal(false);
                    }}
                    style={{
                      border: 'none',
                      backgroundColor: 'transparent',
                      color: 'var(--text-secondary)',
                      cursor: 'pointer',
                      fontSize: '22px',
                      lineHeight: 1
                    }}
                  >
                    &times;
                  </button>
                </div>

                <div style={{ padding: '12px 16px', borderBottom: '1px solid var(--border-color)', flexShrink: 0 }}>
                  <input
                    type="text"
                    placeholder={language === 'uz' ? "Qidirish..." : "Поиск..."}
                    value={modalSearchQuery}
                    onChange={(e) => setModalSearchQuery(e.target.value)}
                    style={{
                      width: '100%',
                      padding: '10px 12px',
                      borderRadius: '6px',
                      border: '1px solid var(--border-color)',
                      backgroundColor: 'var(--bg-primary)',
                      color: 'var(--text-primary)',
                      fontSize: '13px',
                      boxSizing: 'border-box'
                    }}
                  />
                </div>

                <div style={{
                  padding: '12px 16px',
                  overflowY: 'auto',
                  flexGrow: 1,
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '8px'
                }}>
                  {agentProducts
                    .filter(p => {
                      const name = (p.productName || p.name || '').toLowerCase();
                      return name.includes(modalSearchQuery.toLowerCase());
                    })
                    .map(prod => {
                      const cartItem = cashierCart.find(item => item.productId === prod.productId || item.productId === prod.id);
                      const inCartQty = cartItem ? cartItem.quantity : 0;
                      const liveRemaining = (prod.remainingQty !== undefined ? prod.remainingQty : prod.qty) - inCartQty;
                      
                      const isChecked = selectedModalProducts.includes(prod.productId || prod.id);
                      const isOutOfStock = liveRemaining <= 0;

                      return (
                        <div
                          key={prod.id || prod.productId}
                          onClick={() => {
                            if (isOutOfStock) return;
                            const pId = prod.productId || prod.id;
                            if (isChecked) {
                              setSelectedModalProducts(prev => prev.filter(id => id !== pId));
                            } else {
                              setSelectedModalProducts(prev => [...prev, pId]);
                            }
                          }}
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: '12px',
                            padding: '12px',
                            borderRadius: '8px',
                            border: '1px solid var(--border-color)',
                            backgroundColor: isChecked ? 'rgba(13, 148, 136, 0.05)' : 'var(--bg-secondary)',
                            cursor: isOutOfStock ? 'not-allowed' : 'pointer',
                            opacity: isOutOfStock ? 0.5 : 1,
                            transition: 'all 0.2s ease',
                            textAlign: 'left'
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={isChecked}
                            disabled={isOutOfStock}
                            onChange={() => {}}
                            style={{
                              width: '18px',
                              height: '18px',
                              accentColor: 'var(--accent-color)',
                              cursor: isOutOfStock ? 'not-allowed' : 'pointer'
                            }}
                          />

                          <div style={{ display: 'flex', flexDirection: 'column', gap: '2px', flexGrow: 1 }}>
                            <span style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-primary)' }}>
                              {prod.productName}
                            </span>
                            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                              <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                                {parseFloat(prod.price || 0).toLocaleString('uz-UZ')} so'm
                              </span>
                              <span style={{ 
                                fontSize: '10px', 
                                fontWeight: '700', 
                                color: isOutOfStock ? 'var(--danger-color)' : 'var(--success-color)',
                                backgroundColor: isOutOfStock ? 'rgba(239, 68, 68, 0.1)' : 'var(--success-light)',
                                padding: '1px 6px',
                                borderRadius: '4px'
                              }}>
                                {isOutOfStock 
                                  ? (language === 'uz' ? "Tugagan" : "Закончился")
                                  : (language === 'uz' ? `Qoldiq: ${liveRemaining} ${prod.unit || 'dona'}` : `Ост: ${liveRemaining} ${prod.unit || 'шт'}`)
                                }
                              </span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                </div>

                <div style={{
                  padding: '16px 20px',
                  borderTop: '1px solid var(--border-color)',
                  display: 'flex',
                  gap: '10px',
                  flexShrink: 0,
                  backgroundColor: 'var(--bg-secondary)'
                }}>
                  <button
                    onClick={() => {
                      setSelectedModalProducts([]);
                      setShowProductModal(false);
                    }}
                    style={{
                      flexGrow: 1,
                      padding: '12px',
                      border: '1px solid var(--border-color)',
                      backgroundColor: 'var(--bg-primary)',
                      color: 'var(--text-primary)',
                      borderRadius: '8px',
                      fontWeight: '600',
                      cursor: 'pointer',
                      fontSize: '13px'
                    }}
                  >
                    {language === 'uz' ? "Bekor qilish" : "Отмена"}
                  </button>
                  <button
                    onClick={handleAddCheckedProducts}
                    style={{
                      flexGrow: 2,
                      padding: '12px',
                      backgroundColor: 'var(--accent-color)',
                      color: '#fff',
                      border: 'none',
                      borderRadius: '8px',
                      fontWeight: '700',
                      cursor: 'pointer',
                      fontSize: '13px',
                      boxShadow: '0 4px 10px rgba(13, 148, 136, 0.3)'
                    }}
                  >
                    ✓ {language === 'uz' ? `Qo'shish (${selectedModalProducts.length})` : `Добавить (${selectedModalProducts.length})`}
                  </button>
                </div>
              </div>
            </div>
          )}

        </div>
      )}


      {toast && (
        <div style={{
          position: 'fixed',
          top: '24px',
          right: '24px',
          backgroundColor: toast.type === 'error' ? 'var(--danger-color)' : toast.type === 'success' ? 'var(--success-color)' : 'var(--accent-color)',
          color: '#fff',
          padding: '16px 24px',
          borderRadius: '8px',
          boxShadow: 'var(--shadow-lg)',
          zIndex: 99999,
          display: 'flex',
          alignItems: 'center',
          gap: '12px',
          fontWeight: '600',
          fontSize: '14px',
          animation: 'fadeIn 0.2s ease'
        }}>
          <AlertCircle size={18} />
          <span>{toast.message}</span>
          <button 
            onClick={() => setToast(null)} 
            style={{
              background: 'transparent',
              border: 'none',
              color: '#fff',
              cursor: 'pointer',
              fontSize: '16px',
              marginLeft: '8px',
              padding: 0
            }}
          >
            ×
          </button>
        </div>
      )}

      {/* Exit Questionnaire Modal */}
      {showExitQuestionnaire && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.75)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 20000
        }} className="fade-in">
          <div style={{
            width: '90%',
            maxWidth: '380px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '12px',
            padding: '24px',
            boxShadow: 'var(--shadow-lg)',
            display: 'flex',
            flexDirection: 'column',
            gap: '16px'
          }}>
            <h4 style={{ fontSize: '15px', fontWeight: '700', color: 'var(--text-primary)', margin: 0, textAlign: 'center' }}>
              {language === 'uz' ? "Sotuvsiz chiqish sababi" : "Причина выхода без продажи"}
            </h4>
            <p style={{ fontSize: '13px', color: 'var(--text-secondary)', textAlign: 'center', margin: 0 }}>
              {language === 'uz' ? "Nima sababdan do'kondan hech narsa sotmasdan chiqib ketyapsiz?" : "Укажите причину, по которой вы выходите без продажи?"}
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '8px' }}>
              <button
                onClick={() => setExitReason('closed')}
                style={{
                  padding: '12px 16px',
                  borderRadius: '8px',
                  border: '2px solid ' + (exitReason === 'closed' ? 'var(--accent-color)' : 'var(--border-color)'),
                  backgroundColor: exitReason === 'closed' ? 'rgba(13, 148, 136, 0.08)' : 'transparent',
                  color: 'var(--text-primary)',
                  fontSize: '13px',
                  fontWeight: '600',
                  cursor: 'pointer',
                  textAlign: 'left',
                  transition: 'all 0.2s'
                }}
              >
                🚪 {language === 'uz' ? "Do'kon yopiq" : "Магазин закрыт"}
              </button>

              <button
                onClick={() => setExitReason('no_buy')}
                style={{
                  padding: '12px 16px',
                  borderRadius: '8px',
                  border: '2px solid ' + (exitReason === 'no_buy' ? 'var(--accent-color)' : 'var(--border-color)'),
                  backgroundColor: exitReason === 'no_buy' ? 'rgba(13, 148, 136, 0.08)' : 'transparent',
                  color: 'var(--text-primary)',
                  fontSize: '13px',
                  fontWeight: '600',
                  cursor: 'pointer',
                  textAlign: 'left',
                  transition: 'all 0.2s'
                }}
              >
                📦 {language === 'uz' ? "Mahsulot olmadi" : "Товар не куплен"}
              </button>
            </div>

            <div style={{ display: 'flex', gap: '10px', marginTop: '12px' }}>
              <button
                onClick={() => {
                  setShowExitQuestionnaire(false);
                  setExitReason('');
                }}
                style={{
                  flexGrow: 1,
                  padding: '10px',
                  borderRadius: '6px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: 'transparent',
                  color: 'var(--text-secondary)',
                  fontWeight: '600',
                  cursor: 'pointer',
                  fontSize: '13px'
                }}
              >
                {language === 'uz' ? "Orqaga" : "Назад"}
              </button>
              <button
                disabled={!exitReason}
                onClick={() => {
                  const newVisit = {
                    storeId: activeCashierStore.id,
                    storeName: activeCashierStore.storeName,
                    status: 'empty',
                    reason: exitReason === 'closed' ? (language === 'uz' ? "Do'kon yopiq" : "Магазин закрыт") : (language === 'uz' ? "Mahsulot olmadi" : "Товар не куплен"),
                    items: [],
                    date: new Date().toISOString().split('T')[0],
                    time: new Date().toLocaleTimeString('uz-UZ', { hour: '2-digit', minute: '2-digit' })
                  };
                  setVisitedStores(prev => {
                    const updated = [newVisit, ...prev];
                    localStorage.setItem('visited_stores', JSON.stringify(updated));
                    return updated;
                  });

                  // Reset cashier states and close
                  setActiveCashierStore(null);
                  setCashierCart([]);
                  setExitReason('');
                  setShowExitQuestionnaire(false);
                  showAlert(language === 'uz' ? "Hisobot saqlandi va kassa yopildi!" : "Отчет сохранен и касса закрыта!", 'success');
                }}
                style={{
                  flexGrow: 1,
                  padding: '10px',
                  borderRadius: '6px',
                  border: 'none',
                  backgroundColor: exitReason ? 'var(--accent-color)' : 'var(--bg-tertiary)',
                  color: exitReason ? '#fff' : 'var(--text-muted)',
                  fontWeight: '600',
                  cursor: exitReason ? 'pointer' : 'not-allowed',
                  fontSize: '13px'
                }}
              >
                {language === 'uz' ? "Tasdiqlash" : "Подтвердить"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Cashier-specific Terminal Config Modal */}
      {showCashierTerminalConfig && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.75)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 20000
        }} className="fade-in">
          <div style={{
            width: '90%',
            maxWidth: '380px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '12px',
            padding: '24px',
            boxShadow: 'var(--shadow-lg)',
            display: 'flex',
            flexDirection: 'column',
            gap: '16px'
          }}>
            <h4 style={{ fontSize: '16px', fontWeight: '700', color: 'var(--text-primary)', margin: 0 }}>
              {language === 'uz' ? "Terminal Sozlamalari" : "Настройки терминала"}
            </h4>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div>
                <label style={{ fontSize: '11px', color: 'var(--text-secondary)', display: 'block', marginBottom: '4px' }}>
                  {language === 'uz' ? "Terminal IP manzili (Host:Port)" : "IP-адрес терминала (Host:Port)"}
                </label>
                <input 
                  type="text" 
                  value={tindaTerminalIp}
                  onChange={(e) => setTindaTerminalIp(e.target.value)}
                  placeholder="Masalan: 192.168.1.100:8080"
                  style={{ width: '100%', padding: '8px 10px', borderRadius: '6px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600', boxSizing: 'border-box' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', color: 'var(--text-secondary)', display: 'block', marginBottom: '4px' }}>
                  {language === 'uz' ? "Xodim ismi (Terminaldagi login)" : "Имя сотрудника (Логин на терминале)"}
                </label>
                <input 
                  type="text" 
                  value={tindaTerminalLogin}
                  onChange={(e) => setTindaTerminalLogin(e.target.value)}
                  placeholder="Masalan: Xodim Nomi"
                  style={{ width: '100%', padding: '8px 10px', borderRadius: '6px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600', boxSizing: 'border-box' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', color: 'var(--text-secondary)', display: 'block', marginBottom: '4px' }}>
                  {language === 'uz' ? "Terminal PIN kodi" : "ПИН-код терминала"}
                </label>
                <input 
                  type="password" 
                  value={tindaTerminalPin}
                  onChange={(e) => setTindaTerminalPin(e.target.value)}
                  placeholder="Masalan: 1111"
                  style={{ width: '100%', padding: '8px 10px', borderRadius: '6px', border: '1px solid var(--border-color)', backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', fontSize: '13px', fontWeight: '600', boxSizing: 'border-box' }}
                />
              </div>
            </div>

            <div style={{ display: 'flex', gap: '10px', marginTop: '8px' }}>
              <button
                onClick={() => setShowCashierTerminalConfig(false)}
                style={{
                  flexGrow: 1,
                  padding: '10px',
                  borderRadius: '6px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: 'transparent',
                  color: 'var(--text-secondary)',
                  fontWeight: '600',
                  cursor: 'pointer',
                  fontSize: '13px'
                }}
              >
                {language === 'uz' ? "Yopish" : "Закрыть"}
              </button>
              <button
                onClick={() => {
                  localStorage.setItem('tinda_terminal_ip', tindaTerminalIp);
                  localStorage.setItem('tinda_terminal_login', tindaTerminalLogin);
                  localStorage.setItem('tinda_terminal_pin', tindaTerminalPin);
                  setShowCashierTerminalConfig(false);
                  showAlert(language === 'uz' ? "Terminal sozlamalari saqlandi!" : "Настройки терминала сохранены!", 'success');
                }}
                style={{
                  flexGrow: 1,
                  padding: '10px',
                  borderRadius: '6px',
                  border: 'none',
                  backgroundColor: 'var(--accent-color)',
                  color: '#fff',
                  fontWeight: '600',
                  cursor: 'pointer',
                  fontSize: '13px'
                }}
              >
                {language === 'uz' ? "Saqlash" : "Сохранить"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Tinda Terminal Payment Overlay Modal */}
      {tindaPaymentStatus && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          backdropFilter: 'blur(6px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 99999
        }} className="fade-in">
          <div style={{
            width: '90%',
            maxWidth: '440px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '16px',
            padding: '32px 24px',
            boxShadow: 'var(--shadow-xl)',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            textAlign: 'center',
            gap: '24px',
            position: 'relative'
          }}>
            {/* Status Icon & Loader */}
            <div style={{ position: 'relative', width: '80px', height: '80px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              {['connecting', 'logging_in', 'waiting_card'].includes(tindaPaymentStatus) && (
                <div className="tinda-pulse-ring" style={{
                  position: 'absolute',
                  width: '100%',
                  height: '100%',
                  borderRadius: '50%',
                  border: '3px solid var(--accent-color)',
                  animation: 'pulse 1.5s infinite ease-in-out'
                }} />
              )}
              <div style={{
                width: '64px',
                height: '64px',
                borderRadius: '50%',
                backgroundColor: tindaPaymentStatus === 'success' ? 'rgba(34, 197, 94, 0.15)' : tindaPaymentStatus === 'error' ? 'rgba(239, 68, 68, 0.15)' : 'rgba(13, 148, 136, 0.15)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: tindaPaymentStatus === 'success' ? '#22c55e' : tindaPaymentStatus === 'error' ? '#ef4444' : 'var(--accent-color)',
                zIndex: 1
              }}>
                {tindaPaymentStatus === 'success' && <CheckCircle size={32} />}
                {tindaPaymentStatus === 'error' && <AlertCircle size={32} />}
                {['connecting', 'logging_in'].includes(tindaPaymentStatus) && <Loader size={32} className="animate-spin" />}
                {tindaPaymentStatus === 'waiting_card' && <CreditCard size={32} />}
              </div>
            </div>

            {/* Texts */}
            <div>
              <h4 style={{ fontSize: '18px', fontWeight: '700', color: 'var(--text-primary)', marginBottom: '8px' }}>
                {tindaPaymentStatus === 'connecting' && (language === 'uz' ? "Terminalga ulanish..." : "Подключение к терминалу...")}
                {tindaPaymentStatus === 'logging_in' && (language === 'uz' ? "Tizimga kirish..." : "Авторизация...")}
                {tindaPaymentStatus === 'waiting_card' && (language === 'uz' ? "Karta kutilmoqda..." : "Ожидание карты...")}
                {tindaPaymentStatus === 'success' && (language === 'uz' ? "To'lov muvaffaqiyatli o'tdi!" : "Оплата успешно прошла!")}
                {tindaPaymentStatus === 'error' && (language === 'uz' ? "To'lovda xatolik!" : "Ошибка оплаты!")}
              </h4>
              <p style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                {tindaPaymentStatus === 'connecting' && (language === 'uz' ? "Terminal bilan aloqa o'rnatilmoqda, iltimos kuting." : "Устанавливается связь с терминалом, пожалуйста, подождите.")}
                {tindaPaymentStatus === 'logging_in' && (language === 'uz' ? "Terminalda xodim smenasi tekshirilmoqda." : "Проверяется смена сотрудника на терминале.")}
                {tindaPaymentStatus === 'waiting_card' && (language === 'uz' ? "Iltimos, kartani terminalga kiriting yoki tekkizing." : "Пожалуйста, вставьте или приложите карту к терминалу.")}
                {tindaPaymentStatus === 'success' && (language === 'uz' ? "Fiskal chek chop etilmoqda va sotuv yozilmoqda." : "Печатается фискальный чек и записывается продажа.")}
                {tindaPaymentStatus === 'error' && (tindaErrorMessage || (language === 'uz' ? "Noma'lum xatolik yuz berdi." : "Произошла неизвестная ошибка."))}
              </p>
            </div>

            {/* Sum details */}
            <div style={{
              width: '100%',
              backgroundColor: 'var(--bg-primary)',
              borderRadius: '8px',
              padding: '12px 16px',
              border: '1px solid var(--border-color)',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center'
            }}>
              <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                {language === 'uz' ? "To'lov summasi" : "Сумма оплаты"}
              </span>
              <span style={{ fontSize: '16px', fontWeight: '800', color: 'var(--accent-color)' }}>
                {(() => {
                  const subtotal = cashierCart.reduce((acc, item) => acc + (item.price * item.quantity), 0);
                  const discountPercentage = customDiscountEnabled && customDiscountInput ? parseFloat(customDiscountInput) || 0 : cashierDiscount;
                  const discountAmount = Math.round(subtotal * (discountPercentage / 100));
                  const finalTotal = subtotal - discountAmount;
                  return finalTotal.toLocaleString('uz-UZ');
                })()} so'm
              </span>
            </div>

            {/* Action Button */}
            <button
              onClick={() => {
                if (tindaSocket) {
                  tindaSocket.close();
                }
                setTindaPaymentStatus(null);
              }}
              style={{
                width: '100%',
                padding: '12px',
                borderRadius: '8px',
                backgroundColor: tindaPaymentStatus === 'error' ? 'var(--danger-color)' : 'transparent',
                color: tindaPaymentStatus === 'error' ? '#fff' : 'var(--text-secondary)',
                fontWeight: '600',
                fontSize: '13px',
                cursor: 'pointer',
                transition: 'all 0.2s ease',
                border: tindaPaymentStatus === 'error' ? 'none' : '1px solid var(--border-color)'
              }}
            >
              {tindaPaymentStatus === 'error' 
                ? (language === 'uz' ? "Yopish va qayta urinish" : "Закрыть и повторить") 
                : (language === 'uz' ? "Bekor qilish" : "Отмена")}
            </button>
          </div>
        </div>
      )}

      {/* Custom React Confirm Dialog Modal */}
      {confirmDialog && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          backgroundColor: 'rgba(15, 23, 42, 0.75)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 99999
        }} className="fade-in">
          <div style={{
            width: '400px',
            backgroundColor: 'var(--bg-secondary)',
            border: '1px solid var(--border-color)',
            borderRadius: '12px',
            padding: '24px',
            boxShadow: 'var(--shadow-lg)',
            display: 'flex',
            flexDirection: 'column',
            gap: '16px'
          }}>
            <h4 style={{ fontSize: '16px', fontWeight: '600', color: 'var(--text-primary)' }}>
              {language === 'uz' ? 'Tasdiqlash' : 'Подтверждение'}
            </h4>
            <p style={{ fontSize: '14px', color: 'var(--text-secondary)' }}>
              {confirmDialog.message}
            </p>
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end', marginTop: '8px' }}>
              <button
                onClick={confirmDialog.onCancel}
                style={{
                  padding: '8px 16px',
                  borderRadius: '6px',
                  border: '1px solid var(--border-color)',
                  backgroundColor: 'transparent',
                  color: 'var(--text-secondary)',
                  fontWeight: '600',
                  cursor: 'pointer'
                }}
              >
                {t('close')}
              </button>
              <button
                onClick={confirmDialog.onConfirm}
                style={{
                  padding: '8px 16px',
                  borderRadius: '6px',
                  border: 'none',
                  backgroundColor: 'var(--danger-color)',
                  color: '#fff',
                  fontWeight: '600',
                  cursor: 'pointer'
                }}
              >
                {language === 'uz' ? 'Tasdiqlash' : 'Подтвердить'}
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

export default App;
