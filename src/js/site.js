/* Keeps the floating "Call us" button out of the way until the visitor has
   scrolled past the hero. Without JavaScript the button simply stays visible,
   which is the safe fallback. */
(function () {
  'use strict';

  var float = document.querySelector('.call-float');
  if (!float) return;

  float.classList.add('call-float--managed');

  var hero = document.querySelector('.hero');
  var trigger = hero ? hero.offsetTop + hero.offsetHeight - 120 : 300;
  var ticking = false;

  function update() {
    ticking = false;
    var y = window.pageYOffset || document.documentElement.scrollTop;
    float.classList.toggle('is-visible', y > trigger);
  }

  function onScroll() {
    if (ticking) return;
    ticking = true;
    window.requestAnimationFrame(update);
  }

  window.addEventListener('scroll', onScroll, { passive: true });
  window.addEventListener('resize', function () {
    trigger = hero ? hero.offsetTop + hero.offsetHeight - 120 : 300;
    onScroll();
  });
  update();
})();
