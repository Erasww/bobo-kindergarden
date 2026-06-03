// Sticky nav
window.addEventListener('scroll', () => {
  document.getElementById('navbar').classList.toggle('scrolled', window.scrollY > 20);
});

// Mobile menu
function toggleMenu() {
  document.getElementById('navLinks').classList.toggle('open');
}
document.querySelectorAll('.nav-links a').forEach((a) => a.addEventListener('click', () => {
  document.getElementById('navLinks').classList.remove('open');
}));

// FAQ
function toggleFaq(el) {
  const item = el.parentElement;
  document.querySelectorAll('.faq-item').forEach((i) => {
    if (i !== item) {
      i.classList.remove('open');
    }
  });
  item.classList.toggle('open');
}

// Gallery filter
function filterGallery(btn, cat) {
  document.querySelectorAll('.gf-btn').forEach((b) => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.gallery-item').forEach((item) => {
    const show = cat === 'all' || item.dataset.cat === cat;
    item.style.opacity = show ? '1' : '0.2';
    item.style.transform = show ? '' : 'scale(0.95)';
    item.style.transition = 'opacity .3s, transform .3s';
  });
}

async function submitForm(event) {
  if (event) {
    event.preventDefault();
  }

  const form = document.getElementById('enrollForm');
  if (!form.reportValidity()) {
    return;
  }

  const submitButton = document.getElementById('submitButton');
  const successBox = document.getElementById('formSuccess');
  const errorBox = document.getElementById('formError');
  const formData = new FormData(form);
  const data = Object.fromEntries(formData.entries());
  const backendOrigin = 'http://localhost:3000';
  const enrollUrl = `${backendOrigin}/enroll`;

  errorBox.textContent = '';
  submitButton.disabled = true;
  submitButton.textContent = 'Submitting...';

  try {
    const response = await fetch(enrollUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });

    const rawBody = await response.text();
    let result = {};

    if (rawBody) {
      try {
        result = JSON.parse(rawBody);
      } catch (parseError) {
        throw new Error('The website did not receive a valid server response. Start the app with npm start and open http://localhost:3000.');
      }
    }

    if (!response.ok || !result.success) {
      const serverMessage = typeof result.message === 'string' && result.message.trim()
        ? result.message
        : '';
      const fallbackMessage = rawBody
        ? `Request failed with status ${response.status}: ${rawBody}`
        : `Request failed with status ${response.status}.`;

      throw new Error(serverMessage || fallbackMessage);
    }

    form.reset();
    successBox.style.display = 'block';
  } catch (error) {
    console.error('Enrollment error:', error);
    if (error instanceof TypeError) {
      errorBox.textContent = `Cannot reach the server at ${backendOrigin}. Start it with npm start.`;
    } else {
      errorBox.textContent = error.message || 'Something went wrong. Please try again.';
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = '🌟 Submit Enrollment Inquiry';
  }
}

document.getElementById('enrollForm').addEventListener('submit', submitForm);
document.getElementById('submitButton').addEventListener('click', submitForm);

// Schedule item click
document.querySelectorAll('.schedule-item').forEach((item) => {
  item.addEventListener('click', () => {
    document.querySelectorAll('.schedule-item').forEach((i) => i.classList.remove('active'));
    item.classList.add('active');
  });
});

// Reveal on scroll
const reveals = document.querySelectorAll('.reveal');
const observer = new IntersectionObserver((entries) => {
  entries.forEach((e) => {
    if (e.isIntersecting) {
      e.target.classList.add('visible');
      observer.unobserve(e.target);
    }
  });
}, { threshold: 0.1 });
reveals.forEach((r) => observer.observe(r));
