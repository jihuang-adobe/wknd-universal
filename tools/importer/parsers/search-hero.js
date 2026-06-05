/* eslint-disable */
/* global WebImporter */

/**
 * Parser for search-hero
 * Base block: search
 * Source: https://www.bd.com/en-us
 * Selector: .bd-hero-search
 * Generated: 2026-06-05
 *
 * UE Model fields:
 *   - index (text/string) - the search index path
 *   - classes (multiselect) - SKIPPED per hinting rules
 *
 * Source structure:
 *   .bd-hero-search > .bd-hero-search__parent > section > .bd-hero-search__cnt
 *     > .bd-hero-search__wrapper > .bd-hero-search__cnt--form
 *       > form#documentation-other-search
 *         > input[type="search"] (placeholder text)
 *         > button[type="submit"] (SEARCH)
 */
export default function parse(element, { document }) {
  // Extract the search input placeholder as a content indicator for the search index
  const searchInput = element.querySelector('input[type="search"], input[name="search-input"], input[placeholder]');
  const placeholder = searchInput ? (searchInput.getAttribute('placeholder') || 'search') : 'search';

  // The search-hero model has a single field: "index" (text/string)
  // Use the placeholder text as the index value representing search configuration
  const cells = [];

  // Row 1: index field - search configuration path/label
  // xwalk field hint for Universal Editor integration
  const indexCell = document.createElement('div');
  indexCell.appendChild(document.createComment(' field:index '));
  const indexParagraph = document.createElement('p');
  indexParagraph.textContent = placeholder;
  indexCell.appendChild(indexParagraph);
  cells.push([indexCell]);

  const block = WebImporter.Blocks.createBlock(document, { name: 'search-hero', cells });
  element.replaceWith(block);
}
