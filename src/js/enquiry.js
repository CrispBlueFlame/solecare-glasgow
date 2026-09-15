/* Enquiry form.
   Without JavaScript the form still works: it posts normally to Netlify and
   Netlify shows its own thank-you page. With JavaScript we post in the
   background so the confirmation can appear in place, without a page reload. */
(function () {
  'use strict';

  var form = document.getElementById('enquiry-form');
  var confirm = document.getElementById('enquiry-confirm');
  var error = document.getElementById('enquiry-error');
  if (!form || !confirm) return;

  // Older browsers without fetch fall back to a normal form post.
  if (typeof window.fetch !== 'function' || typeof window.FormData !== 'function') return;

  function reveal(el) {
    el.hidden = false;
    var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    try {
      el.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'center' });
    } catch (e) {
      el.scrollIntoView();
    }
  }

  form.addEventListener('submit', function (e) {
    e.preventDefault();

    var button = form.querySelector('button[type="submit"]');
    var original = button ? button.textContent : '';
    if (button) {
      button.disabled = true;
      button.textContent = 'Sending...';
    }
    if (error) error.hidden = true;

    var body = new URLSearchParams(new FormData(form)).toString();

    fetch(form.getAttribute('action') || '/contact/', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body
    })
      .then(function (res) {
        if (!res.ok) throw new Error('Status ' + res.status);
        form.hidden = true;
        reveal(confirm);
        confirm.focus();
      })
      .catch(function () {
        if (button) {
          button.disabled = false;
          button.textContent = original;
        }
        if (error) {
          error.hidden = false;
          reveal(error);
        }
      });
  });
})();
