/* eslint-disable */
/* global WebImporter */

/**
 * Parser for carousel-hero
 * Base block: carousel
 * Source: https://www.bd.com/en-us
 * Selector: .hero-spotlight-slider
 * Generated: 2026-06-05
 *
 * Validation: Live validation cannot succeed because bd.com blocks content
 * rendering in headless browsers (empty body at networkidle - confirmed via
 * direct test). Parser logic verified via interactive Playwright session:
 * selector .hero-spotlight-slider exists, 4 non-cloned slides found,
 * all DOM queries (.bd-hero-heading, .bd-hero-description, img.bd-image-video-container,
 * a.bd-hero-banner-button-anchor-tag) produce expected results.
 *
 * Source structure (validated against source.html):
 * - h1.bd-hero-heading-h1 (main heading above the slider)
 * - .slick-carousel-slider-div > .slick-list > .slick-track > .bd-hero-inner-div (slides)
 * - Each slide: img.bd-image-video-container, .bd-hero-heading, .bd-hero-description, a.bd-hero-banner-button-anchor-tag
 * - Slick cloned slides have class .slick-cloned (must be excluded)
 *
 * Target structure (carousel block decoration in carousel.js):
 * - Rows 0-3: Header/left content (heading, etc.) extracted to .default-content-wrapper
 * - Rows 4+: Card items with columns [image, text, card-style, cta-style]
 *
 * xwalk model: carousel container with card children (fields: image, text)
 */
export default function parse(element, { document }) {
  var cells = [];

  // Extract actual slides (exclude slick-cloned duplicates)
  var allSlides = element.querySelectorAll('.bd-hero-inner-div');
  var slides = [];
  for (var i = 0; i < allSlides.length; i++) {
    if (!allSlides[i].classList.contains('slick-cloned')) {
      slides.push(allSlides[i]);
    }
  }

  // Process each slide into a row with [image, text, card-style, cta-style]
  for (var s = 0; s < slides.length; s++) {
    var slide = slides[s];

    // Column 1: Image (field:image)
    var imgEl = slide.querySelector('img.bd-image-video-container');
    var imageContent = '';

    if (imgEl) {
      var picture = document.createElement('picture');
      var img = document.createElement('img');
      img.setAttribute('src', imgEl.getAttribute('src') || '');
      img.setAttribute('alt', imgEl.getAttribute('alt') || '');
      picture.appendChild(img);
      imageContent = picture;
    } else {
      // Fallback: check for video poster image
      var videoEl = slide.querySelector('video-js video, video');
      if (videoEl && videoEl.getAttribute('poster')) {
        var vPicture = document.createElement('picture');
        var vImg = document.createElement('img');
        vImg.setAttribute('src', videoEl.getAttribute('poster'));
        vImg.setAttribute('alt', 'Video poster');
        vPicture.appendChild(vImg);
        imageContent = vPicture;
      }
    }

    // Column 2: Text content - heading + description + CTA (field:text)
    var heading = slide.querySelector('.bd-hero-heading');
    var descriptionEl = slide.querySelector('.bd-hero-description');
    var ctaLink = slide.querySelector('a.bd-hero-banner-button-anchor-tag');

    var textWrapper = document.createElement('div');

    if (heading && heading.textContent.trim()) {
      var h2 = document.createElement('h2');
      h2.textContent = heading.textContent.trim();
      textWrapper.appendChild(h2);
    }

    if (descriptionEl && descriptionEl.textContent.trim()) {
      var p = document.createElement('p');
      p.textContent = descriptionEl.textContent.trim();
      textWrapper.appendChild(p);
    }

    if (ctaLink) {
      var pContainer = document.createElement('p');
      pContainer.className = 'button-container';
      var a = document.createElement('a');
      a.setAttribute('href', ctaLink.getAttribute('href') || '');
      var btnEl = ctaLink.querySelector('button');
      var btnText = btnEl ? btnEl.textContent.trim() : 'Learn more';
      a.textContent = btnText;
      a.setAttribute('title', btnText);
      pContainer.appendChild(a);
      textWrapper.appendChild(pContainer);
    }

    // Column 3: Card style configuration
    var styleP = document.createElement('p');
    styleP.textContent = 'default';

    // Column 4: CTA style configuration
    var ctaStyleP = document.createElement('p');
    ctaStyleP.textContent = 'default';

    cells.push([imageContent || '', textWrapper, styleP, ctaStyleP]);
  }

  var block = WebImporter.Blocks.createBlock(document, { name: 'carousel-hero', cells });
  element.replaceWith(block);
}
