const wrapper = document.querySelector('.page');
const leftZone = document.querySelector('.left');
const rightZone = document.querySelector('.right');
const slides = document.querySelectorAll('.container');

const scrollAmount = window.innerWidth;

rightZone.addEventListener('click', () => {
  wrapper.scrollLeft += scrollAmount;
});

leftZone.addEventListener('click', () => {
  wrapper.scrollLeft -= scrollAmount;
});

let touchStartX = 0;
let touchEndX = 0;

wrapper.addEventListener('scroll', () => {
  const containerWidth = wrapper.offsetWidth;
  const scrollLeft = wrapper.scrollLeft;

  slides.forEach((slide, index) => {
    const slideOffset = slide.offsetLeft;
    const distanceFromCenter = Math.abs(scrollLeft - slideOffset);
    let opacity = 1 - (distanceFromCenter / containerWidth);
    slide.style.opacity = Math.max(0, Math.min(1, opacity));
  });
});