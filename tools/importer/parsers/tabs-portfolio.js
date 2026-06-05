/* eslint-disable */
/* global WebImporter */

/**
 * Parser: tabs-portfolio
 * Base block: tabs
 * Source: https://www.bd.com/en-us
 * Selector: .bd-image-map.bd-home-page__protfolio
 * Generated: 2026-06-05
 *
 * Extracts a tabbed portfolio section with healthcare setting tabs.
 * Each tab has: icon, label, challenge text, solution text, and a CTA link.
 * Maps to the Tabs container block model (tabs-item children).
 *
 * UE Model (tabs-item fields per row):
 *   Column 1: title (tab label text)
 *   Column 2: content_heading, content_image, content_richtext (grouped by content_ prefix)
 *   Collapsed: content_headingType (Type suffix - not hinted)
 *
 * Validated selectors against source.html:
 *   - ul.hotspot-container__tabs > li (tab labels with img + h2)
 *   - .hotspot-container__tabs-content > div.content (tab panels with description + link)
 *   - .bd-image-map__description-1 (challenge text)
 *   - .bd-image-map__description-2 (solution text)
 *   - a.bd-image-map__link (CTA link)
 */
export default function parse(element, { document }) {
  // Extract tab labels from the tab navigation list
  var tabLabels = element.querySelectorAll('ul.hotspot-container__tabs > li');
  if (!tabLabels.length) {
    tabLabels = element.querySelectorAll('.hotspot-container__tabs li');
  }

  // Extract tab content panels
  var tabPanels = element.querySelectorAll('.hotspot-container__tabs-content > div.content');
  if (!tabPanels.length) {
    tabPanels = element.querySelectorAll('.hotspot-container__tabs-content > div[ids]');
  }

  var cells = [];
  var count = Math.min(tabLabels.length, tabPanels.length);

  for (var i = 0; i < count; i++) {
    var label = tabLabels[i];
    var panel = tabPanels[i];

    // --- Column 1: title (tab label) ---
    var tabHeading = label.querySelector('h2, h3, h4');
    var tabTitle = tabHeading ? tabHeading.textContent.trim() : label.textContent.trim();

    var titleFrag = document.createDocumentFragment();
    titleFrag.appendChild(document.createComment(' field:title '));
    var titleP = document.createElement('p');
    titleP.textContent = tabTitle;
    titleFrag.appendChild(titleP);

    // --- Column 2: grouped content_* fields ---
    var contentFrag = document.createDocumentFragment();

    // content_heading: heading text as h3 (matches model default headingType)
    contentFrag.appendChild(document.createComment(' field:content_heading '));
    var headingEl = document.createElement('h3');
    headingEl.textContent = tabTitle;
    contentFrag.appendChild(headingEl);

    // content_image: icon image from the tab label
    var iconImg = label.querySelector('img');
    if (iconImg) {
      contentFrag.appendChild(document.createComment(' field:content_image '));
      var picture = document.createElement('picture');
      var img = document.createElement('img');
      img.src = iconImg.getAttribute('src') || '';
      img.alt = tabTitle || iconImg.getAttribute('alt') || '';
      picture.appendChild(img);
      contentFrag.appendChild(picture);
    }

    // content_richtext: challenge text + solution text + CTA link
    contentFrag.appendChild(document.createComment(' field:content_richtext '));
    var richtextContainer = document.createElement('div');

    // Challenge section
    var challengeDiv = panel.querySelector('.bd-image-map__description-1, .image-map__description');
    if (challengeDiv) {
      var challengePs = challengeDiv.querySelectorAll('p');
      for (var c = 0; c < challengePs.length; c++) {
        var clonedC = document.createElement('p');
        clonedC.innerHTML = challengePs[c].innerHTML;
        richtextContainer.appendChild(clonedC);
      }
    }

    // Solution section
    var solutionDiv = panel.querySelector('.bd-image-map__description-2');
    if (solutionDiv) {
      var solutionPs = solutionDiv.querySelectorAll('p');
      for (var s = 0; s < solutionPs.length; s++) {
        var clonedS = document.createElement('p');
        clonedS.innerHTML = solutionPs[s].innerHTML;
        richtextContainer.appendChild(clonedS);
      }
    }

    // CTA link (main navigation link, not the "Read More" toggle)
    var ctaLink = panel.querySelector('a.bd-image-map__link');
    if (ctaLink) {
      var link = document.createElement('a');
      link.href = ctaLink.getAttribute('href') || '#';
      // Extract text without embedded images
      var linkText = '';
      for (var n = 0; n < ctaLink.childNodes.length; n++) {
        if (ctaLink.childNodes[n].nodeType === 3) {
          linkText += ctaLink.childNodes[n].textContent;
        }
      }
      link.textContent = linkText.trim() || ctaLink.textContent.replace(/\s+/g, ' ').trim();
      var target = ctaLink.getAttribute('target');
      if (target) {
        link.setAttribute('target', target);
      }
      var linkP = document.createElement('p');
      linkP.appendChild(link);
      richtextContainer.appendChild(linkP);
    }

    contentFrag.appendChild(richtextContainer);

    // Each row: [title column, content column]
    cells.push([titleFrag, contentFrag]);
  }

  var block = WebImporter.Blocks.createBlock(document, { name: 'tabs-portfolio', cells });
  element.replaceWith(block);
}
