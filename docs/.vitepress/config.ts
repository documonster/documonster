import { defineConfig } from "vitepress";

const docsBase = process.env.DOCS_BASE ?? "/";

export default defineConfig({
  // NOTE:
  // - For a custom domain (e.g. https://excelts.dev): set DOCS_BASE='/'
  // - For GitHub project pages (https://<user>.github.io/<repo>/): set DOCS_BASE='/<repo>/'
  base: docsBase,

  locales: {
    root: {
      label: "English",
      lang: "en-US",
      title: "ExcelTS",
      description: "Modern TypeScript Excel workbook manager (XLSX/CSV) for Node.js and browsers.",
      themeConfig: {
        nav: [
          { text: "Guide", link: "/guide/getting-started" },
          { text: "API", link: "/api/" },
          { text: "Examples", link: "/examples" },
          { text: "GitHub", link: "https://github.com/cjnoname/excelts" }
        ],
        sidebar: {
          "/guide/": [
            {
              text: "Guide",
              items: [
                { text: "Getting Started", link: "/guide/getting-started" },
                { text: "Browser Support", link: "/guide/browser" },
                { text: "Streaming", link: "/guide/streaming" }
              ]
            }
          ],
          "/api/": [
            {
              text: "API",
              items: [{ text: "Overview", link: "/api/" }]
            }
          ],
          "/": [
            {
              text: "Reference",
              items: [
                { text: "README", link: "/reference/readme" },
                { text: "README (中文)", link: "/reference/readme-zh" },
                { text: "Changelog", link: "/reference/changelog" }
              ]
            }
          ]
        }
      }
    },
    zh: {
      label: "简体中文",
      lang: "zh-CN",
      link: "/zh/",
      title: "ExcelTS",
      description: "现代 TypeScript Excel 工作簿管理库（XLSX/CSV）。",
      themeConfig: {
        nav: [
          { text: "指南", link: "/zh/guide/getting-started" },
          { text: "API", link: "/zh/api/" },
          { text: "示例", link: "/examples" },
          { text: "GitHub", link: "https://github.com/cjnoname/excelts" }
        ],
        sidebar: {
          "/zh/guide/": [
            {
              text: "指南",
              items: [
                { text: "快速开始", link: "/zh/guide/getting-started" },
                { text: "浏览器支持", link: "/zh/guide/browser" },
                { text: "流式读写", link: "/zh/guide/streaming" }
              ]
            }
          ],
          "/zh/api/": [
            {
              text: "API",
              items: [{ text: "概览", link: "/zh/api/" }]
            }
          ],
          "/": [
            {
              text: "参考",
              items: [
                { text: "README", link: "/reference/readme" },
                { text: "README（中文）", link: "/reference/readme-zh" },
                { text: "Changelog", link: "/reference/changelog" }
              ]
            }
          ]
        }
      }
    }
  },

  themeConfig: {
    socialLinks: [{ icon: "github", link: "https://github.com/cjnoname/excelts" }],
    search: { provider: "local" },
    footer: {
      message: "Released under the MIT License.",
      copyright: `Copyright © ${new Date().getFullYear()} cjnoname`
    }
  }
});
