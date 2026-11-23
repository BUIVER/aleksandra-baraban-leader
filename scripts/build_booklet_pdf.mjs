#!/usr/bin/env node
/**
 * Экспорт booklet.html в PDF для печати (A4, без UI).
 * Запуск из корня репозитория: node scripts/build_booklet_pdf.mjs
 */
import { chromium } from "playwright";
import { fileURLToPath } from "url";
import path from "path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const htmlPath = path.join(root, "booklet.html");
const outPath = path.join(root, "Буклет_Лидеры_Года_Александра_Барабан.pdf");
const fileUrl = `file://${htmlPath}`;

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  await page.goto(fileUrl, { waitUntil: "networkidle" });
  await page.pdf({
    path: outPath,
    format: "A4",
    printBackground: true,
    preferCSSPageSize: true,
    margin: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  console.log("PDF:", outPath);
} finally {
  await browser.close();
}