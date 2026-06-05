/* eslint-disable */
var CustomImportScript = (() => {
  var __defProp = Object.defineProperty;
  var __defProps = Object.defineProperties;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __getOwnPropDescs = Object.getOwnPropertyDescriptors;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __getOwnPropSymbols = Object.getOwnPropertySymbols;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __propIsEnum = Object.prototype.propertyIsEnumerable;
  var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
  var __spreadValues = (a, b) => {
    for (var prop in b || (b = {}))
      if (__hasOwnProp.call(b, prop))
        __defNormalProp(a, prop, b[prop]);
    if (__getOwnPropSymbols)
      for (var prop of __getOwnPropSymbols(b)) {
        if (__propIsEnum.call(b, prop))
          __defNormalProp(a, prop, b[prop]);
      }
    return a;
  };
  var __spreadProps = (a, b) => __defProps(a, __getOwnPropDescs(b));
  var __export = (target, all) => {
    for (var name in all)
      __defProp(target, name, { get: all[name], enumerable: true });
  };
  var __copyProps = (to, from, except, desc) => {
    if (from && typeof from === "object" || typeof from === "function") {
      for (let key of __getOwnPropNames(from))
        if (!__hasOwnProp.call(to, key) && key !== except)
          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
    }
    return to;
  };
  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

  // tools/importer/import-homepage.js
  var import_homepage_exports = {};
  __export(import_homepage_exports, {
    default: () => import_homepage_default
  });

  // tools/importer/parsers/carousel-hero.js
  function parse(element, { document }) {
    var cells = [];
    var allSlides = element.querySelectorAll(".bd-hero-inner-div");
    var slides = [];
    for (var i = 0; i < allSlides.length; i++) {
      if (!allSlides[i].classList.contains("slick-cloned")) {
        slides.push(allSlides[i]);
      }
    }
    for (var s = 0; s < slides.length; s++) {
      var slide = slides[s];
      var imgEl = slide.querySelector("img.bd-image-video-container");
      var imageContent = "";
      if (imgEl) {
        var picture = document.createElement("picture");
        var img = document.createElement("img");
        img.setAttribute("src", imgEl.getAttribute("src") || "");
        img.setAttribute("alt", imgEl.getAttribute("alt") || "");
        picture.appendChild(img);
        imageContent = picture;
      } else {
        var videoEl = slide.querySelector("video-js video, video");
        if (videoEl && videoEl.getAttribute("poster")) {
          var vPicture = document.createElement("picture");
          var vImg = document.createElement("img");
          vImg.setAttribute("src", videoEl.getAttribute("poster"));
          vImg.setAttribute("alt", "Video poster");
          vPicture.appendChild(vImg);
          imageContent = vPicture;
        }
      }
      var heading = slide.querySelector(".bd-hero-heading");
      var descriptionEl = slide.querySelector(".bd-hero-description");
      var ctaLink = slide.querySelector("a.bd-hero-banner-button-anchor-tag");
      var textWrapper = document.createElement("div");
      if (heading && heading.textContent.trim()) {
        var h2 = document.createElement("h2");
        h2.textContent = heading.textContent.trim();
        textWrapper.appendChild(h2);
      }
      if (descriptionEl && descriptionEl.textContent.trim()) {
        var p = document.createElement("p");
        p.textContent = descriptionEl.textContent.trim();
        textWrapper.appendChild(p);
      }
      if (ctaLink) {
        var pContainer = document.createElement("p");
        pContainer.className = "button-container";
        var a = document.createElement("a");
        a.setAttribute("href", ctaLink.getAttribute("href") || "");
        var btnEl = ctaLink.querySelector("button");
        var btnText = btnEl ? btnEl.textContent.trim() : "Learn more";
        a.textContent = btnText;
        a.setAttribute("title", btnText);
        pContainer.appendChild(a);
        textWrapper.appendChild(pContainer);
      }
      var styleP = document.createElement("p");
      styleP.textContent = "default";
      var ctaStyleP = document.createElement("p");
      ctaStyleP.textContent = "default";
      cells.push([imageContent || "", textWrapper, styleP, ctaStyleP]);
    }
    var block = WebImporter.Blocks.createBlock(document, { name: "carousel-hero", cells });
    element.replaceWith(block);
  }

  // tools/importer/parsers/search-hero.js
  function parse2(element, { document }) {
    const searchInput = element.querySelector('input[type="search"], input[name="search-input"], input[placeholder]');
    const placeholder = searchInput ? searchInput.getAttribute("placeholder") || "search" : "search";
    const cells = [];
    const indexCell = document.createElement("div");
    indexCell.appendChild(document.createComment(" field:index "));
    const indexParagraph = document.createElement("p");
    indexParagraph.textContent = placeholder;
    indexCell.appendChild(indexParagraph);
    cells.push([indexCell]);
    const block = WebImporter.Blocks.createBlock(document, { name: "search-hero", cells });
    element.replaceWith(block);
  }

  // tools/importer/parsers/cards-news.js
  function parse3(element, { document }) {
    const items = element.querySelectorAll("li");
    const cells = [];
    items.forEach((item) => {
      const img = item.querySelector("img");
      const imageFrag = document.createDocumentFragment();
      imageFrag.appendChild(document.createComment(" field:image "));
      if (img) {
        const pic = document.createElement("img");
        pic.src = img.getAttribute("src") || "";
        pic.alt = img.getAttribute("alt") || "";
        imageFrag.appendChild(pic);
      }
      const textFrag = document.createDocumentFragment();
      textFrag.appendChild(document.createComment(" field:text "));
      const categoryP = item.querySelector(":scope > p");
      if (categoryP && categoryP.textContent.trim()) {
        const catEl = document.createElement("p");
        catEl.textContent = categoryP.textContent.trim();
        textFrag.appendChild(catEl);
      }
      const titleLink = item.querySelector("a.title") || item.querySelector("a");
      if (titleLink) {
        const h3 = document.createElement("h3");
        const a = document.createElement("a");
        a.href = titleLink.getAttribute("href") || "";
        a.textContent = titleLink.textContent.trim();
        const target = titleLink.getAttribute("target");
        if (target) {
          a.setAttribute("target", target);
        }
        h3.appendChild(a);
        textFrag.appendChild(h3);
      }
      const desc = item.querySelector(".description") || item.querySelector(":scope > div");
      if (desc) {
        const innerP = desc.querySelector("p");
        const descEl = document.createElement("p");
        descEl.textContent = (innerP || desc).textContent.trim();
        textFrag.appendChild(descEl);
      }
      cells.push([imageFrag, textFrag]);
    });
    const block = WebImporter.Blocks.createBlock(document, { name: "cards-news", cells });
    element.replaceWith(block);
  }

  // tools/importer/parsers/tabs-portfolio.js
  function parse4(element, { document }) {
    var tabLabels = element.querySelectorAll("ul.hotspot-container__tabs > li");
    if (!tabLabels.length) {
      tabLabels = element.querySelectorAll(".hotspot-container__tabs li");
    }
    var tabPanels = element.querySelectorAll(".hotspot-container__tabs-content > div.content");
    if (!tabPanels.length) {
      tabPanels = element.querySelectorAll(".hotspot-container__tabs-content > div[ids]");
    }
    var cells = [];
    var count = Math.min(tabLabels.length, tabPanels.length);
    for (var i = 0; i < count; i++) {
      var label = tabLabels[i];
      var panel = tabPanels[i];
      var tabHeading = label.querySelector("h2, h3, h4");
      var tabTitle = tabHeading ? tabHeading.textContent.trim() : label.textContent.trim();
      var titleFrag = document.createDocumentFragment();
      titleFrag.appendChild(document.createComment(" field:title "));
      var titleP = document.createElement("p");
      titleP.textContent = tabTitle;
      titleFrag.appendChild(titleP);
      var contentFrag = document.createDocumentFragment();
      contentFrag.appendChild(document.createComment(" field:content_heading "));
      var headingEl = document.createElement("h3");
      headingEl.textContent = tabTitle;
      contentFrag.appendChild(headingEl);
      var iconImg = label.querySelector("img");
      if (iconImg) {
        contentFrag.appendChild(document.createComment(" field:content_image "));
        var picture = document.createElement("picture");
        var img = document.createElement("img");
        img.src = iconImg.getAttribute("src") || "";
        img.alt = tabTitle || iconImg.getAttribute("alt") || "";
        picture.appendChild(img);
        contentFrag.appendChild(picture);
      }
      contentFrag.appendChild(document.createComment(" field:content_richtext "));
      var richtextContainer = document.createElement("div");
      var challengeDiv = panel.querySelector(".bd-image-map__description-1, .image-map__description");
      if (challengeDiv) {
        var challengePs = challengeDiv.querySelectorAll("p");
        for (var c = 0; c < challengePs.length; c++) {
          var clonedC = document.createElement("p");
          clonedC.innerHTML = challengePs[c].innerHTML;
          richtextContainer.appendChild(clonedC);
        }
      }
      var solutionDiv = panel.querySelector(".bd-image-map__description-2");
      if (solutionDiv) {
        var solutionPs = solutionDiv.querySelectorAll("p");
        for (var s = 0; s < solutionPs.length; s++) {
          var clonedS = document.createElement("p");
          clonedS.innerHTML = solutionPs[s].innerHTML;
          richtextContainer.appendChild(clonedS);
        }
      }
      var ctaLink = panel.querySelector("a.bd-image-map__link");
      if (ctaLink) {
        var link = document.createElement("a");
        link.href = ctaLink.getAttribute("href") || "#";
        var linkText = "";
        for (var n = 0; n < ctaLink.childNodes.length; n++) {
          if (ctaLink.childNodes[n].nodeType === 3) {
            linkText += ctaLink.childNodes[n].textContent;
          }
        }
        link.textContent = linkText.trim() || ctaLink.textContent.replace(/\s+/g, " ").trim();
        var target = ctaLink.getAttribute("target");
        if (target) {
          link.setAttribute("target", target);
        }
        var linkP = document.createElement("p");
        linkP.appendChild(link);
        richtextContainer.appendChild(linkP);
      }
      contentFrag.appendChild(richtextContainer);
      cells.push([titleFrag, contentFrag]);
    }
    var block = WebImporter.Blocks.createBlock(document, { name: "tabs-portfolio", cells });
    element.replaceWith(block);
  }

  // tools/importer/parsers/columns-careers.js
  function parse5(element, { document }) {
    const image = element.querySelector(".bd-image-card__image-container img, .bd-image-card__image");
    const categoryEl = element.querySelector(".bd-image-card__category");
    const headingEl = element.querySelector(".bd-image-card__heading");
    const detailsEl = element.querySelector(".bd-image-card__details");
    const ctaLink = element.querySelector("a.bd-image-card__link-learn-more, .bd-image-card__content a");
    const leftCol = [];
    if (image) {
      leftCol.push(image);
    }
    const rightCol = [];
    if (categoryEl) {
      const eyebrowP = categoryEl.querySelector("p");
      if (eyebrowP) {
        rightCol.push(eyebrowP);
      } else {
        rightCol.push(categoryEl);
      }
    }
    if (headingEl) {
      const headingP = headingEl.querySelector("p");
      if (headingP) {
        const h2 = document.createElement("h2");
        h2.textContent = headingP.textContent;
        rightCol.push(h2);
      } else {
        rightCol.push(headingEl);
      }
    }
    if (detailsEl) {
      const paragraphs = detailsEl.querySelectorAll("p");
      paragraphs.forEach((p) => {
        rightCol.push(p);
      });
    }
    if (ctaLink) {
      rightCol.push(ctaLink);
    }
    const cells = [
      [leftCol, rightCol]
    ];
    const block = WebImporter.Blocks.createBlock(document, { name: "columns-careers", cells });
    element.replaceWith(block);
  }

  // tools/importer/transformers/bd-cleanup.js
  var TransformHook = { beforeTransform: "beforeTransform", afterTransform: "afterTransform" };
  function transform(hookName, element, payload) {
    if (hookName === TransformHook.beforeTransform) {
      const { document } = payload;
      WebImporter.DOMUtils.remove(element, [
        "header",
        "footer",
        "nav",
        ".header",
        ".footer",
        ".navigation",
        "script",
        "style",
        "noscript",
        "iframe",
        "link",
        ".slick-cloned",
        "video-js",
        ".slider-controls-max-width",
        ".bd-contact-us",
        ".onetrust-consent-sdk",
        "#onetrust-consent-sdk",
        ".userway-s3",
        '[class*="cookie"]',
        '[class*="consent"]',
        '[id*="cookie"]',
        '[id*="consent"]'
      ]);
      const main = document.querySelector("main") || document.querySelector("#content-div") || document.querySelector(".cmp-container");
      if (main && main !== element) {
        while (element.firstChild) {
          element.removeChild(element.firstChild);
        }
        element.appendChild(main);
      }
    }
    if (hookName === TransformHook.afterTransform) {
      WebImporter.DOMUtils.remove(element, ["iframe", "link", "noscript", "script", "style"]);
      element.querySelectorAll("*").forEach((el) => {
        el.removeAttribute("tabindex");
        el.removeAttribute("aria-hidden");
        el.removeAttribute("role");
        el.removeAttribute("aria-describedby");
        el.removeAttribute("aria-label");
        el.removeAttribute("aria-controls");
        el.removeAttribute("aria-live");
        el.removeAttribute("aria-disabled");
        el.removeAttribute("aria-checked");
        el.removeAttribute("aria-expanded");
        el.removeAttribute("aria-haspopup");
        el.removeAttribute("aria-valuemin");
        el.removeAttribute("aria-valuemax");
        el.removeAttribute("aria-valuenow");
        el.removeAttribute("aria-valuetext");
        el.removeAttribute("aria-selected");
        el.removeAttribute("aria-atomic");
        el.removeAttribute("translate");
        el.removeAttribute("dir");
      });
    }
  }

  // tools/importer/transformers/bd-sections.js
  var TransformHook2 = { beforeTransform: "beforeTransform", afterTransform: "afterTransform" };
  function transform2(hookName, element, payload) {
    if (hookName === TransformHook2.afterTransform) {
      const template = payload && payload.template;
      if (!template) return;
      const sections = template.sections;
      if (!sections || sections.length < 2) return;
      const document = element.ownerDocument;
      for (let i = sections.length - 1; i >= 0; i--) {
        const section = sections[i];
        if (!section.selector) continue;
        const sectionEl = element.querySelector(section.selector);
        if (!sectionEl) continue;
        if (section.style) {
          const block = WebImporter.Blocks.createBlock(document, {
            name: "Section Metadata",
            cells: [["style", section.style]]
          });
          sectionEl.after(block);
        }
        if (i > 0) {
          const hr = document.createElement("hr");
          sectionEl.before(hr);
        }
      }
    }
  }

  // tools/importer/import-homepage.js
  var parsers = {
    "carousel-hero": parse,
    "search-hero": parse2,
    "cards-news": parse3,
    "tabs-portfolio": parse4,
    "columns-careers": parse5
  };
  var PAGE_TEMPLATE = {
    name: "homepage",
    description: "BD US homepage with hero, product categories, news, and corporate information",
    urls: [
      "https://www.bd.com/en-us"
    ],
    blocks: [
      {
        name: "carousel-hero",
        instances: [".hero-spotlight-slider"]
      },
      {
        name: "search-hero",
        instances: [".bd-hero-search"]
      },
      {
        name: "cards-news",
        instances: [".bd-trending"]
      },
      {
        name: "tabs-portfolio",
        instances: [".bd-image-map.bd-home-page__protfolio"]
      },
      {
        name: "columns-careers",
        instances: [".consolidated-image-card"]
      }
    ],
    sections: [
      {
        id: "section-1",
        name: "Hero Carousel",
        selector: ".hero-spotlight-slider",
        style: null,
        blocks: ["carousel-hero"],
        defaultContent: []
      },
      {
        id: "section-2",
        name: "Search Bar",
        selector: ".bd-hero-search",
        style: null,
        blocks: ["search-hero"],
        defaultContent: []
      },
      {
        id: "section-3",
        name: "News Trending",
        selector: ".bd-trending",
        style: null,
        blocks: ["cards-news"],
        defaultContent: []
      },
      {
        id: "section-4",
        name: "Our Portfolio",
        selector: ".bd-container.bd-container__grey",
        style: "grey",
        blocks: ["tabs-portfolio"],
        defaultContent: [".bd-image-map__section-name", ".bd-image-map__section-description"]
      },
      {
        id: "section-5",
        name: "Careers",
        selector: ".consolidated-image-card",
        style: "dark",
        blocks: ["columns-careers"],
        defaultContent: []
      }
    ]
  };
  var transformers = [
    transform,
    ...PAGE_TEMPLATE.sections && PAGE_TEMPLATE.sections.length > 1 ? [transform2] : []
  ];
  function executeTransformers(hookName, element, payload) {
    const enhancedPayload = __spreadProps(__spreadValues({}, payload), {
      template: PAGE_TEMPLATE
    });
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
            section: blockDef.section || null
          });
        });
      });
    });
    console.log(`Found ${pageBlocks.length} block instances on page`);
    return pageBlocks;
  }
  var import_homepage_default = {
    transform: (payload) => {
      const { document, url, html, params } = payload;
      const main = document.body;
      executeTransformers("beforeTransform", main, payload);
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
      executeTransformers("afterTransform", main, payload);
      const hr = document.createElement("hr");
      main.appendChild(hr);
      WebImporter.rules.createMetadata(main, document);
      WebImporter.rules.transformBackgroundImages(main, document);
      WebImporter.rules.adjustImageUrls(main, url, params.originalURL);
      const path = WebImporter.FileUtils.sanitizePath(
        new URL(params.originalURL).pathname.replace(/\/$/, "").replace(/\.html$/, "")
      );
      return [{
        element: main,
        path,
        report: {
          title: document.title,
          template: PAGE_TEMPLATE.name,
          blocks: pageBlocks.map((b) => b.name)
        }
      }];
    }
  };
  return __toCommonJS(import_homepage_exports);
})();
