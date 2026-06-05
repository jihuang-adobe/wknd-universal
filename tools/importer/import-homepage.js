/* eslint-disable */
/* global WebImporter */

import carouselHeroParser from './parsers/carousel-hero.js';
import searchHeroParser from './parsers/search-hero.js';
import cardsNewsParser from './parsers/cards-news.js';
import tabsPortfolioParser from './parsers/tabs-portfolio.js';
import columnsCareersParser from './parsers/columns-careers.js';

import bdCleanupTransformer from './transformers/bd-cleanup.js';
import bdSectionsTransformer from './transformers/bd-sections.js';

const parsers = {
  'carousel-hero': carouselHeroParser,
  'search-hero': searchHeroParser,
  'cards-news': cardsNewsParser,
  'tabs-portfolio': tabsPortfolioParser,
  'columns-careers': columnsCareersParser,
};

const PAGE_TEMPLATE = {
  name: 'homepage',
  description: 'BD US homepage with hero, product categories, news, and corporate information',
  urls: [
    'https://www.bd.com/en-us',
  ],
  blocks: [
    {
      name: 'carousel-hero',
      instances: ['.hero-spotlight-slider'],
    },
    {
      name: 'search-hero',
      instances: ['.bd-hero-search'],
    },
    {
      name: 'cards-news',
      instances: ['.bd-trending'],
    },
    {
      name: 'tabs-portfolio',
      instances: ['.bd-image-map.bd-home-page__protfolio'],
    },
    {
      name: 'columns-careers',
      instances: ['.consolidated-image-card'],
    },
  ],
  sections: [
    {
      id: 'section-1',
      name: 'Hero Carousel',
      selector: '.hero-spotlight-slider',
      style: null,
      blocks: ['carousel-hero'],
      defaultContent: [],
    },
    {
      id: 'section-2',
      name: 'Search Bar',
      selector: '.bd-hero-search',
      style: null,
      blocks: ['search-hero'],
      defaultContent: [],
    },
    {
      id: 'section-3',
      name: 'News Trending',
      selector: '.bd-trending',
      style: null,
      blocks: ['cards-news'],
      defaultContent: [],
    },
    {
      id: 'section-4',
      name: 'Our Portfolio',
      selector: '.bd-container.bd-container__grey',
      style: 'grey',
      blocks: ['tabs-portfolio'],
      defaultContent: ['.bd-image-map__section-name', '.bd-image-map__section-description'],
    },
    {
      id: 'section-5',
      name: 'Careers',
      selector: '.consolidated-image-card',
      style: 'dark',
      blocks: ['columns-careers'],
      defaultContent: [],
    },
  ],
};

const transformers = [
  bdCleanupTransformer,
  ...(PAGE_TEMPLATE.sections && PAGE_TEMPLATE.sections.length > 1 ? [bdSectionsTransformer] : []),
];

function executeTransformers(hookName, element, payload) {
  const enhancedPayload = {
    ...payload,
    template: PAGE_TEMPLATE,
  };

  transformers.forEach((transformerFn) => {
    try {
      transformerFn.call(null, hookName, element, enhancedPayload);
    } catch (e) {
      console.error(`Transformer failed at ${hookName}:`, e);
    }
  });
}

function findBlocksOnPage(document, template) {
  const pageBlocks = [];

  template.blocks.forEach((blockDef) => {
    blockDef.instances.forEach((selector) => {
      const elements = document.querySelectorAll(selector);
      if (elements.length === 0) {
        console.warn(`Block "${blockDef.name}" selector not found: ${selector}`);
      }
      elements.forEach((element) => {
        pageBlocks.push({
          name: blockDef.name,
          selector,
          element,
          section: blockDef.section || null,
        });
      });
    });
  });

  console.log(`Found ${pageBlocks.length} block instances on page`);
  return pageBlocks;
}

export default {
  transform: (payload) => {
    const { document, url, html, params } = payload;

    const main = document.body;

    executeTransformers('beforeTransform', main, payload);

    const pageBlocks = findBlocksOnPage(document, PAGE_TEMPLATE);

    pageBlocks.forEach((block) => {
      const parser = parsers[block.name];
      if (parser) {
        try {
          parser(block.element, { document, url, params });
        } catch (e) {
          console.error(`Failed to parse ${block.name} (${block.selector}):`, e);
        }
      } else {
        console.warn(`No parser found for block: ${block.name}`);
      }
    });

    executeTransformers('afterTransform', main, payload);

    const hr = document.createElement('hr');
    main.appendChild(hr);
    WebImporter.rules.createMetadata(main, document);
    WebImporter.rules.transformBackgroundImages(main, document);
    WebImporter.rules.adjustImageUrls(main, url, params.originalURL);

    const path = WebImporter.FileUtils.sanitizePath(
      new URL(params.originalURL).pathname.replace(/\/$/, '').replace(/\.html$/, ''),
    );

    return [{
      element: main,
      path,
      report: {
        title: document.title,
        template: PAGE_TEMPLATE.name,
        blocks: pageBlocks.map((b) => b.name),
      },
    }];
  },
};
