/* eslint-disable */
/* global WebImporter */

/**
 * Transformer: BD site-wide cleanup.
 * Removes non-authorable content from BD.com pages.
 * All selectors validated against captured DOM from migration-work/cleaned.html.
 */
const TransformHook = { beforeTransform: 'beforeTransform', afterTransform: 'afterTransform' };

export default function transform(hookName, element, payload) {
  if (hookName === TransformHook.beforeTransform) {
    const { document } = payload;

    WebImporter.DOMUtils.remove(element, [
      'header',
      'footer',
      'nav',
      '.header',
      '.footer',
      '.navigation',
      'script',
      'style',
      'noscript',
      'iframe',
      'link',
      '.slick-cloned',
      'video-js',
      '.slider-controls-max-width',
      '.bd-contact-us',
      '.onetrust-consent-sdk',
      '#onetrust-consent-sdk',
      '.userway-s3',
      '[class*="cookie"]',
      '[class*="consent"]',
      '[id*="cookie"]',
      '[id*="consent"]',
    ]);

    // Scope to main content area if available
    const main = document.querySelector('main') || document.querySelector('#content-div') || document.querySelector('.cmp-container');
    if (main && main !== element) {
      while (element.firstChild) {
        element.removeChild(element.firstChild);
      }
      element.appendChild(main);
    }
  }

  if (hookName === TransformHook.afterTransform) {
    WebImporter.DOMUtils.remove(element, ['iframe', 'link', 'noscript', 'script', 'style']);

    element.querySelectorAll('*').forEach((el) => {
      el.removeAttribute('tabindex');
      el.removeAttribute('aria-hidden');
      el.removeAttribute('role');
      el.removeAttribute('aria-describedby');
      el.removeAttribute('aria-label');
      el.removeAttribute('aria-controls');
      el.removeAttribute('aria-live');
      el.removeAttribute('aria-disabled');
      el.removeAttribute('aria-checked');
      el.removeAttribute('aria-expanded');
      el.removeAttribute('aria-haspopup');
      el.removeAttribute('aria-valuemin');
      el.removeAttribute('aria-valuemax');
      el.removeAttribute('aria-valuenow');
      el.removeAttribute('aria-valuetext');
      el.removeAttribute('aria-selected');
      el.removeAttribute('aria-atomic');
      el.removeAttribute('translate');
      el.removeAttribute('dir');
    });
  }
}
