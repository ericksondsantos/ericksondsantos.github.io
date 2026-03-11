const wrapper = document.querySelector('.slider-wrapper');
const leftZone = document.querySelector('.left-zone');
const rightZone = document.querySelector('.right-zone');
const slides = document.querySelectorAll('.slide');

const scrollAmount = window.innerWidth;

// Click Logic
rightZone.addEventListener('click', () => {
  wrapper.scrollLeft += scrollAmount;
});

leftZone.addEventListener('click', () => {
  wrapper.scrollLeft -= scrollAmount;
});

// Swipe Logic
let touchStartX = 0;
let touchEndX = 0;

wrapper.addEventListener('touchstart', e => {
  touchStartX = e.changedTouches[0].screenX;
});

wrapper.addEventListener('touchend', e => {
  touchEndX = e.changedTouches[0].screenX;
  handleSwipe();
});

wrapper.addEventListener('scroll', () => {
  const containerWidth = wrapper.offsetWidth;
  const scrollLeft = wrapper.scrollLeft;

  slides.forEach((slide, index) => {
    // Calculate how close the slide is to the center of the view
    const slideOffset = slide.offsetLeft;
    const distanceFromCenter = Math.abs(scrollLeft - slideOffset);
    
    // Calculate opacity (1 when centered, 0 when far away)
    // Adjust the '1000' value to change how fast the fade happens
    let opacity = 1 - (distanceFromCenter / containerWidth);
    
    // Clamp the value between 0 and 1
    slide.style.opacity = Math.max(0, Math.min(1, opacity));
  });
});

function handleSwipe() {
  const threshold = 50; // Minimum pixels to count as a swipe
  if (touchStartX - touchEndX > threshold) {
    // Swiped Left (move right)
    wrapper.scrollLeft += scrollAmount;
  } else if (touchEndX - touchStartX > threshold) {
    // Swiped Right (move left)
    wrapper.scrollLeft -= scrollAmount;
  }
}