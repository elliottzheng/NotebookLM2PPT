import { defineConfig } from 'vitepress'

// https://vitepress.dev/reference/site-config
export default defineConfig({
  srcDir: "docs",
  head: [['link', { rel: 'icon', href: '/favicon.ico' }]],
  title: "NotebookLM2PPT",
  description: "将 PDF 转换为可编辑 PowerPoint 的自动化工具",
  themeConfig: {
    
    // https://vitepress.dev/reference/default-theme-config
    nav: [
      { text: '首页', link: '/' },
      { text: '用户指南', items: [
        { text: '功能介绍', link: '/features' },
        { text: '快速开始', link: '/quickstart' },
        { text: '使用教程', link: '/tutorial' },
        { text: '样例对比', link: '/compare' }
      ]},
      { text: '技术文档', items: [
        { text: '工作原理', link: '/implementation' },
        { text: 'MinerU 优化', link: '/mineru' },
      ]},
      { text: '更新日志', link: '/changelog' }
    ],

    sidebar: [
      {
        text: '用户指南',
        items: [
          { text: '首页', link: '/' },
          { text: '功能介绍', link: '/features' },
          { text: '快速开始', link: '/quickstart' },
          { text: '使用教程', link: '/tutorial' },
          { text: '样例对比', link: '/compare' }
        ]
      },
      {
        text: '技术文档',
        items: [
          { text: '工作原理', link: '/implementation' },
          { text: 'MinerU 优化', link: '/mineru' },
        ]
      },
      {
        text: '其他',
        items: [
          { text: '更新日志', link: '/changelog' }
        ]
      }
    ],

    socialLinks: [
      { icon: 'github', link: 'https://github.com/elliottzheng/NotebookLM2PPT' }
    ],

    logo: '/logo_tiny.png',
    
    footer: {
      message: '基于 MIT 许可证开源',
      copyright: 'Copyright © 2026-Present NotebookLM2PPT'
    },

    search: {
      provider: 'local'
    },

    lastUpdated: {
      text: '最后更新',
      formatOptions: {
        dateStyle: 'short',
        timeStyle: 'medium'
      }
    }
  }
})
