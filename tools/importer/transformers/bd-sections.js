/* eslint-disable */
/* global WebImporter */

/**
 * Transformer: BD section breaks and section metadata.
 * Creates section boundaries (<hr>) and Section Metadata blocks based on template sections.
 * Runs in afterTransform only. Processes sections in reverse order.
 * All selectors from page-templates.json, validated against captured DOM.
 *
 * Sections (from payload.template.sections):
 * 1. Hero Carousel - .hero-spotlight-slider (no style)
 * 2. Search Bar - .bd-hero-search (no style)
 * 3. News Trending - .bd-trending (no style)
 * 4. Our Portfolio - .bd-container.bd-container__grey (style: grey)
 * 5. Careers - .consolidated-image-card (style: dark)
 */
const TransformHook = { beforeTransform: 'beforeTransform', afterTransform: 'afterTransform' };

export default function transform(hookName, element, payload) {
  if (hookName === TransformHook.afterTransform) {
    const template = payload && payload.template;
    if (!template) return;
    const sections = template.sections;
    if (!sections || sections.length < 2) return;

    const document = element.ownerDocument;

    // Process sections in reverse order to avoid DOM position shifts
    for (let i = sections.length - 1; i >= 0; i--) {
      const section = sections[i];
      if (!section.selector) continue;

      const sectionEl = element.querySelector(section.selector);
      if (!sectionEl) continue;

      // Add Section Metadata block after the section element if it has a style
      if (section.style) {
        const block = WebImporter.Blocks.createBlock(document, {
          name: 'Section Metadata',
          cells: [['style', section.style]],
        });
        sectionEl.after(block);
      }

      // Add <hr> before this section element (except for the first section)
      if (i > 0) {
        const hr = document.createElement('hr');
        sectionEl.before(hr);
      }
    }
  }
}
