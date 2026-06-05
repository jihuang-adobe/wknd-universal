/* eslint-disable */
/* global WebImporter */

/**
 * Parser: columns-careers
 * Base block: columns
 * Source: https://www.bd.com/en-us
 * Selector: .consolidated-image-card
 * Description: Two-column layout with image on left and text content (eyebrow, heading, paragraphs, CTA link) on right.
 * Generated: 2026-06-05
 * Note: Live validation may fail if the source page no longer contains this element.
 * All selectors validated against cached source.html from migration-work/block-context/columns-careers/.
 * Selectors verified: .bd-image-card__image, .bd-image-card__category, .bd-image-card__heading, .bd-image-card__details, a.bd-image-card__link-learn-more
 */
export default function parse(element, { document }) {
  // Column 1: Image
  const image = element.querySelector('.bd-image-card__image-container img, .bd-image-card__image');

  // Column 2: Text content
  const categoryEl = element.querySelector('.bd-image-card__category');
  const headingEl = element.querySelector('.bd-image-card__heading');
  const detailsEl = element.querySelector('.bd-image-card__details');
  const ctaLink = element.querySelector('a.bd-image-card__link-learn-more, .bd-image-card__content a');

  // Build left column (image)
  const leftCol = [];
  if (image) {
    leftCol.push(image);
  }

  // Build right column (text content with eyebrow, heading, description, CTA)
  const rightCol = [];

  // Eyebrow / category text
  if (categoryEl) {
    const eyebrowP = categoryEl.querySelector('p');
    if (eyebrowP) {
      rightCol.push(eyebrowP);
    } else {
      rightCol.push(categoryEl);
    }
  }

  // Heading
  if (headingEl) {
    const headingP = headingEl.querySelector('p');
    if (headingP) {
      // Convert to h2 for proper semantic structure
      const h2 = document.createElement('h2');
      h2.textContent = headingP.textContent;
      rightCol.push(h2);
    } else {
      rightCol.push(headingEl);
    }
  }

  // Description paragraphs
  if (detailsEl) {
    const paragraphs = detailsEl.querySelectorAll('p');
    paragraphs.forEach((p) => {
      rightCol.push(p);
    });
  }

  // CTA link
  if (ctaLink) {
    rightCol.push(ctaLink);
  }

  // Build cells: one row with two columns
  const cells = [
    [leftCol, rightCol],
  ];

  const block = WebImporter.Blocks.createBlock(document, { name: 'columns-careers', cells });
  element.replaceWith(block);
}
