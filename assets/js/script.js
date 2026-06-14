
document.addEventListener('DOMContentLoaded', () => {

  // Mobile menu toggle
  const mobileMenuButton = document.getElementById('mobile-menu-button');
  const mobileMenu = document.getElementById('mobile-menu');

  if (mobileMenuButton && mobileMenu) {
    mobileMenuButton.addEventListener('click', function() {
      mobileMenu.classList.toggle('hidden');
    });
  }

  // Smooth scrolling for navigation links
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
      if (this.getAttribute('href').startsWith('#')) {
        e.preventDefault();
        const targetId = this.getAttribute('href').substring(1);
        const targetElement = document.getElementById(targetId);
        
        if (targetElement) {
          targetElement.scrollIntoView({ behavior: 'smooth' });
          if (mobileMenu) mobileMenu.classList.add('hidden');
        }
      }
    });
  });

  // Video Modal Logic
  const modal = document.getElementById('video-modal');
  const modalVideoContainer = document.getElementById('modal-video-container');
  const closeModalButton = document.getElementById('close-modal');

  function openVideoModal(videoSrc, videoType = 'video/mp4') {
    if (!modal) return;
    
    // Clear previous content
    modalVideoContainer.innerHTML = '';
    
    const videoElement = document.createElement('video');
    videoElement.controls = true;
    videoElement.autoplay = true;
    videoElement.className = 'w-full h-full rounded-lg shadow-2xl'; // Tailwind classes
    
    const sourceElement = document.createElement('source');
    sourceElement.src = videoSrc;
    sourceElement.type = videoType;
    
    videoElement.appendChild(sourceElement);
    modalVideoContainer.appendChild(videoElement);
    
    videoElement.innerHTML += 'Votre navigateur ne supporte pas la lecture de vid&eacute;os.';

    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden'; 
  }

  function closeVideoModal() {
    if (!modal) return;
    modal.classList.add('hidden');
    modalVideoContainer.innerHTML = ''; 
    document.body.style.overflow = ''; 
  }

  // Bind click events to demo buttons
  document.querySelectorAll('.demo-button').forEach(button => {
    button.addEventListener('click', (e) => {
      e.preventDefault();
      const videoSrc = button.getAttribute('data-video-src');
      if (videoSrc) {
        openVideoModal(videoSrc);
      } else {
        alert("La vid\u00e9o de d\u00e9mo n'est pas encore disponible pour ce projet.");
      }
    });
  });

  if (closeModalButton) {
    closeModalButton.addEventListener('click', closeVideoModal);
  }

  // Close when clicking strictly outside the content area
  if (modal) {
    modal.addEventListener('click', (e) => {
      // e.target is the element that triggered the event
      // e.currentTarget is the element the listener is attached to (modal background)
      if (e.target === modal) {
        closeVideoModal();
      }
    });
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal && !modal.classList.contains('hidden')) {
      closeVideoModal();
    }
  });

});
