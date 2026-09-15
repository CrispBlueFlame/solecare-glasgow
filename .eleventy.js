module.exports = function (eleventyConfig) {

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
