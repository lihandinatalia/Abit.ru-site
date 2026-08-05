const PptxGenJS = require('pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();
const dir = 'C:\\\\Users\\\\User\\\\Presents y landings';

pptx.layout = 'LAYOUT_16x9';
pptx.author = 'Наталья Лиханди';
pptx.title = 'Предложение по модернизации сайта Abit.ru';
pptx.subject = 'Продажа нового сайта ООО «Абит»';

const colors = {
    bg: '050507',
    card: '0A0B10',
    text: 'F1F2F5',
    muted: '9AA0A8',
    cyan: '00D4FF',
    violet: '7C3AED',
    pink: 'F43F7A',
    border: '1A1C24'
};

function gradientText(slide, text, x, y, w, h, options = {}) {
    const shape = slide.addText(text, Object.assign({
        x, y, w, h,
        fontFace: 'Inter',
        color: colors.cyan,
        bold: true
    }, options));
    return shape;
}

function sectionTitle(slide, title, x, y, w) {
    return slide.addText(title, {
        x, y, w,
        fontFace: 'Inter',
        fontSize: 32,
        color: colors.text,
        bold: true,
        align: 'center'
    });
}

function bodyText(slide, text, x, y, w, options = {}) {
    return slide.addText(text, Object.assign({
        x, y, w,
        fontFace: 'Inter',
        fontSize: 16,
        color: colors.muted,
        align: 'left',
        lineSpacing: 30
    }, options));
}

function addSlideBg(slide) {
    slide.background = { color: colors.bg };
}

function addAccentBar(slide, x, y, w, h, color) {
    slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: color });
}

function addDecorativeLines(slide) {
    slide.addShape(pptx.ShapeType.line, { x: 0.5, y: 0.4, w: 2, h: 0, line: { color: colors.cyan, width: 1.5 } });
    slide.addShape(pptx.ShapeType.line, { x: 12.5, y: 0.4, w: -1.5, h: 0, line: { color: colors.violet, width: 1.5 } });
}

// Slide 1: Title
const s1 = pptx.addSlide();
addSlideBg(s1);
addDecorativeLines(s1);
s1.addText('Предложение по модернизации сайта', { x: 1, y: 2.2, w: 11.33, h: 1, fontFace: 'Inter', fontSize: 42, color: colors.text, bold: true, align: 'center' });
gradientText(s1, 'Abit.ru', 1, 3.15, 11.33, 0.8, { fontSize: 56, align: 'center' });
s1.addText('Современный сайт для IT-компании: хай-тек дизайн, адаптивность, продающая структура', {
    x: 1, y: 4.3, w: 11.33, h: 0.7, fontFace: 'Inter', fontSize: 18, color: colors.muted, align: 'center'
});
s1.addText('ООО «Абит» — IT-сопровождение и инфраструктура под ключ', {
    x: 1, y: 6.5, w: 11.33, h: 0.5, fontFace: 'Inter', fontSize: 16, color: colors.muted, align: 'center'
});

// Slide 2: Current situation
const s2 = pptx.addSlide();
addSlideBg(s2);
sectionTitle(s2, 'Зачем обновлять сайт?', 1, 0.7, 11.33);
const bullets = [
    { title: 'Первое впечатление о компании', text: 'Сайт — основная точка контакта для потенциальных клиентов. Устаревший дизайн снижает доверие.' },
    { title: 'Конкуренция в IT-сопровождении', text: 'Заказчики сравнивают предложения. Современный сайт отличает от конкурентов.' },
    { title: 'Мобильный трафик', text: 'Более 60% посещений — со смартфонов и планшетов. Сайт должен быть удобным на любом экране.' },
    { title: 'Продающая логика', text: 'Структура должна вести посетителя от проблемы к решению и заявке.' }
];
let y = 1.8;
bullets.forEach((b, i) => {
    const cx = (i % 2) * 5.9 + 0.6;
    const cy = Math.floor(i / 2) * 2.5 + 1.7;
    s2.addShape(pptx.ShapeType.roundRect, { x: cx, y: cy, w: 5.5, h: 2.1, fill: colors.card, line: { color: colors.border }, rectRadius: 0.2 });
    s2.addText(b.title, { x: cx + 0.25, y: cy + 0.2, w: 5, h: 0.5, fontFace: 'Inter', fontSize: 18, color: colors.text, bold: true });
    s2.addText(b.text, { x: cx + 0.25, y: cy + 0.75, w: 5, h: 1.2, fontFace: 'Inter', fontSize: 14, color: colors.muted, lineSpacing: 24 });
});

// Slide 3: What we propose
const s3 = pptx.addSlide();
addSlideBg(s3);
sectionTitle(s3, 'Что предлагаем', 1, 0.7, 11.33);
const items = [
    ['Дизайн в стиле хай-тек', 'Тёмная тема, неоновые акценты, стекломорфизм, анимации. Визуальная метафора надёжности и технологичности.'],
    ['Полная адаптивность', 'Оптимизация под десктоп, планшеты и смартфоны. Корректное отображение меню, карточек и форм.'],
    ['Продающая структура', 'Главная, услуги, решения, проекты, контакты. Каждая страница ведёт к обращению.'],
    ['Интерактивные элементы', 'Раскрывающиеся панели цен, анимированные блоки, удобная форма заявки.'],
    ['Скорость и технологичность', 'Лёгкий статический код, быстрая загрузка, удобная дальнейшая поддержка.']
];
y = 1.7;
items.forEach((it, i) => {
    s3.addShape(pptx.ShapeType.roundRect, { x: 0.6, y, w: 12.1, h: 0.95, fill: colors.card, line: { color: colors.border }, rectRadius: 0.15 });
    s3.addText(String(i + 1), { x: 0.85, y: y + 0.23, w: 0.5, h: 0.5, fontFace: 'Inter', fontSize: 20, color: colors.cyan, bold: true, align: 'center' });
    s3.addText(it[0], { x: 1.5, y: y + 0.22, w: 3.3, h: 0.5, fontFace: 'Inter', fontSize: 16, color: colors.text, bold: true });
    s3.addText(it[1], { x: 5.0, y: y + 0.18, w: 7.5, h: 0.6, fontFace: 'Inter', fontSize: 14, color: colors.muted, lineSpacing: 22 });
    y += 1.15;
});

// Slide 4: Structure
const s4 = pptx.addSlide();
addSlideBg(s4);
sectionTitle(s4, 'Структура сайта', 1, 0.7, 11.33);
const pages = [
    ['Главная', 'Hero-блок, преимущества, услуги, импортозамещение, о компании, призыв к действию.'],
    ['Услуги и цены', 'Раскрывающиеся группы услуг с ценами из прайс-листа.'],
    ['Решения', 'Типовые конфигурации: офис, ИБ, облако, импортозамещение.'],
    ['Проекты', 'Кейсы и примеры реализованных задач с тегами.'],
    ['Контакты', 'Телефон, email, регион, режим работы и форма заявки.'],
    ['Политика конфиденциальности', 'Готовая страница для согласия при отправке форм.']
];
y = 1.7;
pages.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = col * 6.15 + 0.5;
    const cy = row * 1.9 + 1.6;
    s4.addShape(pptx.ShapeType.roundRect, { x: cx, y: cy, w: 5.8, h: 1.6, fill: colors.card, line: { color: colors.border }, rectRadius: 0.15 });
    s4.addText(p[0], { x: cx + 0.25, y: cy + 0.2, w: 5.3, h: 0.45, fontFace: 'Inter', fontSize: 17, color: colors.text, bold: true });
    s4.addText(p[1], { x: cx + 0.25, y: cy + 0.65, w: 5.3, h: 0.85, fontFace: 'Inter', fontSize: 13, color: colors.muted, lineSpacing: 20 });
});

// Slide 5: Hero screenshot
const s5 = pptx.addSlide();
addSlideBg(s5);
sectionTitle(s5, 'Главная страница', 1, 0.5, 11.33);
try {
    s5.addImage({ path: path.join(dir, 'slide-hero.png'), x: 0.9, y: 1.3, w: 11.5, h: 5.5 });
} catch (e) {}
s5.addText('Акцентный hero-блок с 3D-визуалом, ключевыми цифрами и призывом к действию', {
    x: 1, y: 7.05, w: 11.33, h: 0.5, fontFace: 'Inter', fontSize: 14, color: colors.muted, align: 'center'
});

// Slide 6: Services screenshot
const s6 = pptx.addSlide();
addSlideBg(s6);
sectionTitle(s6, 'Страница услуг и цен', 1, 0.5, 11.33);
try {
    s6.addImage({ path: path.join(dir, 'slide-services.png'), x: 0.9, y: 1.3, w: 11.5, h: 5.5 });
} catch (e) {}
s6.addText('Интерактивный прайс-лист с раскрывающимися группами услуг и ценами', {
    x: 1, y: 7.05, w: 11.33, h: 0.5, fontFace: 'Inter', fontSize: 14, color: colors.muted, align: 'center'
});

// Slide 7: Solutions and contacts
const s7 = pptx.addSlide();
addSlideBg(s7);
s7.addText('Решения и контакты', { x: 1, y: 0.5, w: 11.33, h: 0.7, fontFace: 'Inter', fontSize: 32, color: colors.text, bold: true, align: 'center' });
try {
    s7.addImage({ path: path.join(dir, 'slide-solutions.png'), x: 0.5, y: 1.4, w: 6.0, h: 3.2 });
    s7.addImage({ path: path.join(dir, 'slide-contacts.png'), x: 6.6, y: 1.4, w: 6.0, h: 3.2 });
} catch (e) {}
s7.addText('Готовые решения для бизнеса', { x: 0.5, y: 4.75, w: 6.0, h: 0.4, fontFace: 'Inter', fontSize: 14, color: colors.cyan, align: 'center', bold: true });
s7.addText('Удобная форма обратной связи', { x: 6.6, y: 4.75, w: 6.0, h: 0.4, fontFace: 'Inter', fontSize: 14, color: colors.cyan, align: 'center', bold: true });

// Slide 8: Mobile
const s8 = pptx.addSlide();
addSlideBg(s8);
sectionTitle(s8, 'Адаптация под мобильные устройства', 1, 0.5, 11.33);
try {
    s8.addImage({ path: path.join(dir, 'slide-mobile.png'), x: 4.7, y: 1.35, w: 3.6, h: 5.5 });
} catch (e) {}
s8.addText('Мобильная версия сохраняет читаемость, контрастность меню и удобные кнопки. Сайт одинаково удобен на десктопе, планшете и смартфоне.', {
    x: 1.2, y: 7.05, w: 10.5, h: 0.6, fontFace: 'Inter', fontSize: 14, color: colors.muted, align: 'center', lineSpacing: 24
});

// Slide 9: Benefits
const s9 = pptx.addSlide();
addSlideBg(s9);
sectionTitle(s9, 'Что получит компания', 1, 0.7, 11.33);
const benefits = [
    ['Рост доверия', 'Современный дизайн повышает восприятие профессионализма и надёжности.'],
    ['Удобство клиентов', 'Посетители быстро находят услуги, цены и контакты.'],
    ['Больше заявок', 'Продающая структура и форма на каждой странице увеличивают конверсию.'],
    ['Готовность к росту', 'Легко добавлять новые услуги, кейсы и страницы.']
];
y = 1.8;
benefits.forEach((b, i) => {
    const cx = (i % 2) * 5.9 + 0.6;
    const cy = Math.floor(i / 2) * 2.5 + 1.7;
    s9.addShape(pptx.ShapeType.roundRect, { x: cx, y: cy, w: 5.5, h: 2.1, fill: colors.card, line: { color: colors.border }, rectRadius: 0.2 });
    s9.addText(b[0], { x: cx + 0.25, y: cy + 0.2, w: 5, h: 0.5, fontFace: 'Inter', fontSize: 18, color: colors.cyan, bold: true });
    s9.addText(b[1], { x: cx + 0.25, y: cy + 0.75, w: 5, h: 1.2, fontFace: 'Inter', fontSize: 15, color: colors.muted, lineSpacing: 24 });
});

// Slide 10: Next steps / contacts
const s10 = pptx.addSlide();
addSlideBg(s10);
sectionTitle(s10, 'Следующие шаги', 1, 0.7, 11.33);
s10.addShape(pptx.ShapeType.roundRect, { x: 1.0, y: 1.7, w: 11.0, h: 3.3, fill: colors.card, line: { color: colors.border }, rectRadius: 0.2 });
s10.addText('Если вас устраивает предложение, жду от вас:', {
    x: 1.3, y: 1.95, w: 10.4, h: 0.5, fontFace: 'Inter', fontSize: 19, color: colors.text, bold: true
});
const needs = [
    'Логотип фирмы',
    'Логины и доступы для публикации сайта',
    'Домен для размещения',
    'Правки и пожелания по содержанию сайта'
];
y = 2.55;
needs.forEach(n => {
    s10.addText('▸', { x: 1.4, y, w: 0.4, h: 0.4, fontFace: 'Inter', fontSize: 16, color: colors.cyan });
    s10.addText(n, { x: 1.9, y, w: 9.5, h: 0.4, fontFace: 'Inter', fontSize: 16, color: colors.muted });
    y += 0.55;
});

s10.addShape(pptx.ShapeType.roundRect, { x: 1.0, y: 5.35, w: 11.0, h: 1.6, fill: '0A0B10', line: { color: colors.border }, rectRadius: 0.2 });
s10.addText('С уважением,', { x: 1.3, y: 5.55, w: 10.4, h: 0.35, fontFace: 'Inter', fontSize: 16, color: colors.text });
s10.addText('Наталья Лиханди', { x: 1.3, y: 5.9, w: 10.4, h: 0.4, fontFace: 'Inter', fontSize: 20, color: colors.cyan, bold: true });
s10.addText('Email: lihandin@mail.ru', { x: 6.5, y: 5.9, w: 5.0, h: 0.35, fontFace: 'Inter', fontSize: 15, color: colors.muted });
s10.addText('Телефон: +7 (914) 895-58-15', { x: 6.5, y: 6.25, w: 5.0, h: 0.35, fontFace: 'Inter', fontSize: 15, color: colors.muted });

pptx.writeFile({ fileName: path.join(dir, 'Предложение по модернизации сайта Abit.pptx') });
