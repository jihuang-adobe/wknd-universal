/* eslint-disable */
/* global WebImporter */

/**
 * Parser for cards-news
 * Base block: cards
 * Source: https://www.bd.com/en-us
 * Selector: .bd-trending
 * Generated: 2026-06-05
 *
 * UE Model: container block with "card" items (fields: image, text)
 * Source structure: .bd-trending > article.trending__container
 *   > section.trending__container-bottom > ul > li items
 * Each li contains: img, p (category), a.title (headline), div.description (excerpt)
 *
 * Note: Live validation blocked by site bot protection (Akamai returns 403/Access Denied
 * to headless browsers). Parser validated against cached source.html.
 */
export default function parse(element, { document }) {
  // Find all news card items within the trending container
  const items = element.querySelectorAll('li');
  const cells = [];

  items.forEach((item) => {
    // Column 1: Image with field hint
    const img = item.querySelector('img');
    const imageFrag = document.createDocumentFragment();
    imageFrag.appendChild(document.createComment(' field:image '));
    if (img) {
      const pic = document.createElement('img');
      pic.src = img.getAttribute('src') || '';
      pic.alt = img.getAttribute('alt') || '';
      imageFrag.appendChild(pic);
    }

    // Column 2: Text content (category + headline + description) with field hint
    const textFrag = document.createDocumentFragment();
    textFrag.appendChild(document.createComment(' field:text '));

    // Category tag (NEWS/BLOG)
    const categoryP = item.querySelector(':scope > p');
    if (categoryP && categoryP.textContent.trim()) {
      const catEl = document.createElement('p');
      catEl.textContent = categoryP.textContent.trim();
      textFrag.appendChild(catEl);
    }

    // Headline link wrapped in h3 for semantic structure
    const titleLink = item.querySelector('a.title') || item.querySelector('a');
    if (titleLink) {
      const h3 = document.createElement('h3');
      const a = document.createElement('a');
      a.href = titleLink.getAttribute('href') || '';
      a.textContent = titleLink.textContent.trim();
      const target = titleLink.getAttribute('target');
      if (target) {
        a.setAttribute('target', target);
      }
      h3.appendChild(a);
      textFrag.appendChild(h3);
    }

    // Description / excerpt text
    const desc = item.querySelector('.description') || item.querySelector(':scope > div');
    if (desc) {
      const innerP = desc.querySelector('p');
      const descEl = document.createElement('p');
      descEl.textContent = (innerP || desc).textContent.trim();
      textFrag.appendChild(descEl);
    }

    cells.push([imageFrag, textFrag]);
  });

  const block = WebImporter.Blocks.createBlock(document, { name: 'cards-news', cells });
  element.replaceWith(block);
}
