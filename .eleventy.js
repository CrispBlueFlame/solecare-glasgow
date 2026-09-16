const { HtmlBasePlugin } = require("@11ty/eleventy");

module.exports = function (eleventyConfig) {

  // Every link in the templates is written from the root, for example
  // /services/. That is correct when the site sits on its own domain, which
  // is where it will end up. The GitHub preview address puts it in a
  // subfolder instead, so this plugin rewrites those links to match whatever
  // --pathprefix the build was given. No prefix given, nothing changes.
  eleventyConfig.addPlugin(HtmlBasePlugin);

  // Copy static assets straight through, untouched by the template engine.
  eleventyConfig.addPassthroughCopy("src/style.css");
  eleventyConfig.addPassthroughCopy("src/favicon.svg");
  eleventyConfig.addPassthroughCopy("src/js");
  eleventyConfig.addPassthroughCopy("src/admin");
  eleventyConfig.addPassthroughCopy("src/images");

  // Strip spaces from a phone number so one CMS field drives both the
  // readable text and the tel: link.
  eleventyConfig.addFilter("telLink", (value) =>
    String(value || "").replace(/[^0-9+]/g, "")
  );

  // WhatsApp needs the number in international form with no symbols, so a
  // UK mobile typed as 07700 000000 has to go out as 447700000000.
  eleventyConfig.addFilter("waLink", (value) => {
    let n = String(value || "").replace(/[^0-9+]/g, "");
    if (n.startsWith("+")) n = n.slice(1);
    if (n.startsWith("0")) n = "44" + n.slice(1);
    return n;
  });

  // The admin panel is copied verbatim, never run through the template engine.
  eleventyConfig.ignores.add("src/admin/**");

  eleventyConfig.addWatchTarget("src/style.css");

  return {
    dir: {
      input: "src",
      output: "_site",
      includes: "_includes",
      data: "_data"
    },
    markdownTemplateEngine: "njk",
    htmlTemplateEngine: "njk",
    templateFormats: ["njk", "md", "html"]
  };
};
