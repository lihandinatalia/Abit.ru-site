(function () {
    'use strict';

    document.body.classList.add('js-enabled');

    const burger = document.querySelector('.burger');
    const nav = document.querySelector('.nav');
    const navLinks = document.querySelectorAll('.nav__link');
    const reveals = document.querySelectorAll('.reveal');
    const form = document.getElementById('contactForm');
    const yearEl = document.getElementById('year');

    // Mobile menu
    if (burger && nav) {
        burger.addEventListener('click', () => {
            const isOpen = nav.classList.toggle('is-open');
            burger.classList.toggle('is-active', isOpen);
            burger.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
        });

        navLinks.forEach(link => {
            link.addEventListener('click', () => {
                nav.classList.remove('is-open');
                burger.classList.remove('is-active');
                burger.setAttribute('aria-expanded', 'false');
            });
        });
    }

    // Accordion: smooth animation for details elements
    document.querySelectorAll('details.service-group').forEach(details => {
        const summary = details.querySelector('summary');
        const content = details.querySelector('.service-group__content');
        let isAnimating = false;

        summary.addEventListener('click', (e) => {
            if (isAnimating) {
                e.preventDefault();
                return;
            }
            e.preventDefault();
            isAnimating = true;

            const isOpen = details.hasAttribute('open');
            const height = content.scrollHeight;

            if (isOpen) {
                content.style.height = height + 'px';
                content.style.overflow = 'hidden';
                requestAnimationFrame(() => {
                    content.style.transition = 'height 0.25s ease';
                    content.style.height = '0px';
                });
                setTimeout(() => {
                    details.removeAttribute('open');
                    content.style.height = '';
                    content.style.overflow = '';
                    content.style.transition = '';
                    isAnimating = false;
                }, 250);
            } else {
                details.setAttribute('open', '');
                content.style.height = '0px';
                content.style.overflow = 'hidden';
                requestAnimationFrame(() => {
                    content.style.transition = 'height 0.25s ease';
                    content.style.height = height + 'px';
                });
                setTimeout(() => {
                    content.style.height = '';
                    content.style.overflow = '';
                    content.style.transition = '';
                    isAnimating = false;
                }, 250);
            }
        });
    });

    // Active nav link on scroll
    const sections = document.querySelectorAll('section[id]');
    function setActiveNav() {
        const scrollY = window.scrollY + 100;
        sections.forEach(section => {
            const top = section.offsetTop;
            const height = section.offsetHeight;
            const id = section.getAttribute('id');
            const link = document.querySelector(`.nav__link[href="#${id}"]`);
            if (link && scrollY >= top && scrollY < top + height) {
                document.querySelectorAll('.nav__link').forEach(l => l.classList.remove('active'));
                link.classList.add('active');
            }
        });
    }

    // Reveal animations
    const revealObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
                revealObserver.unobserve(entry.target);
            }
        });
    }, { threshold: 0, rootMargin: '0px 0px 0px 0px' });

    reveals.forEach(el => revealObserver.observe(el));

    // Fallback: ensure reveals are visible if observer doesn't fire (e.g. screenshots, print)
    setTimeout(() => {
        reveals.forEach(el => {
            if (!el.classList.contains('visible')) {
                el.classList.add('visible');
            }
        });
    }, 2000);

    window.addEventListener('scroll', () => {
        document.body.classList.toggle('scrolled', window.scrollY > 20);
        setActiveNav();
    });

    // Contact form: build mailto body
    if (form) {
        form.addEventListener('submit', (e) => {
            const fd = new FormData(form);
            const name = fd.get('name') || '';
            const company = fd.get('company') || '';
            const contact = fd.get('contact') || '';
            const task = fd.get('task') || '';
            const message = fd.get('message') || '';

            const body = [
                `Имя: ${name}`,
                company ? `Компания: ${company}` : '',
                `Контакт: ${contact}`,
                task ? `Тема: ${task}` : '',
                '',
                'Сообщение:',
                message
            ].filter(Boolean).join('\n');

            form.action = `mailto:info@abit-it.ru?subject=${encodeURIComponent(`Заявка с сайта abit-it.ru от ${name}`)}&body=${encodeURIComponent(body)}`;
        });
    }

    if (yearEl) {
        yearEl.textContent = new Date().getFullYear();
    }
})();
