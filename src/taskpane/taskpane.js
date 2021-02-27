/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "core-js/stable";
import "regenerator-runtime/runtime";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import { TextlintKernel } from "@textlint/kernel";
import { moduleInterop } from "@textlint/module-interop";

const textLint = new TextlintKernel();

const lintOptions = {
  // rulePaths: ["./"],
  ext: ".txt",
  plugins: [
    {
      pluginId: "text",
      plugin: moduleInterop(require("@textlint/textlint-plugin-text"))
    }
  ],
  rules: [
    {
      ruleId: "no-doubled-joshi",
      rule: moduleInterop(require("textlint-rule-no-doubled-joshi")),
      options: {
        min_interval: 1,
        allow: ["も", "や"]
      }
    },
    {
      ruleId: "general-novel-style-ja",
      rule: moduleInterop(require("textlint-rule-general-novel-style-ja")),
      options: {
        chars_leading_paragraph: false
      }
    },
    {
      ruleId: "@textlint-ja/no-synonyms",
      rule: moduleInterop(require("@textlint-ja/textlint-rule-no-synonyms")),
      options: {
        dictUrl: devMode
          ? "/dict/sudachi-synonyms-dictionary.json"
          : "/msword-textlint/dict/sudachi-synonyms-dictionary.json"
      }
    },
    {
      ruleId: "prefer-tari-tari",
      rule: moduleInterop(require("textlint-rule-prefer-tari-tari")),
      options: {}
    },
    {
      ruleId: "@textlint-ja/textlint-rule-no-insert-dropping-sa",
      rule: moduleInterop(require("@textlint-ja/textlint-rule-no-insert-dropping-sa")),
      options: {}
    },
    {
      ruleId: "no-dropping-the-ra",
      rule: moduleInterop(require("textlint-rule-no-dropping-the-ra")),
      options: {}
    },
    {
      ruleId: "no-double-negative-ja",
      rule: moduleInterop(require("textlint-rule-no-double-negative-ja")),
      options: {}
    },
    {
      ruleId: "ja-hiragana-keishikimeishi",
      rule: moduleInterop(require("textlint-rule-ja-hiragana-keishikimeishi")),
      options: {}
    },
    {
      ruleId: "ja-hiragana-fukushi",
      rule: moduleInterop(require("textlint-rule-ja-hiragana-fukushi")),
      options: {
        rulePath: devMode ? "/dict/fukushi.yml" : "/msword-textlint/dict/fukushi.yml"
      }
    }
  ]
};

function lintText(text) {
  return textLint.lintText(text, lintOptions);
}

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    window.kuromojin = {
      dicPath: devMode ? "/dict" : "/msword-textlint/dict"
    };
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("lint").onclick = run;
    document.getElementById("clear").onclick = clear;
  }
});

function proceedText(pi, allP, li, allL) {
  return `パラグラフ: ${pi}/${allP}, 指摘: ${li}/${allL}`;
}

export async function run() {
  await clear();

  return Word.run(async context => {
    const processing = document.getElementById("processing");
    const proceed = document.getElementById("proceed");
    processing.style.visibility = "visible";
    proceed.style.visibility = "visible";
    context.document.body.load("paragraphs");
    await context.sync();
    context.document.body.paragraphs.load("items");
    await context.sync();
    try {
      await context.document.body.paragraphs.items.reduce(async (prev, item, pi) => {
        try {
          prev && (await prev);
          proceed.innerText = proceedText(pi, context.document.body.paragraphs.items.length, 0, 0);
          const results = await lintText(item.text);
          if (results.messages.length === 0) {
            return;
          }
          const charRanges = item.getTextRanges(["\n"]);
          charRanges.load();
          await charRanges.context.sync();
          await results.messages.reduce(async (prev, result, li) => {
            prev && (await prev);
            console.log(result);
            proceed.innerText = proceedText(
              pi,
              context.document.body.paragraphs.items.length,
              li,
              results.messages.length
            );
            const ranges = charRanges.items[result.line - 1]
              .search("*", {
                matchWildcards: true
              })
              .load({
                select: "items",
                skip: result.column - 1,
                top: 1
              });
            await ranges.context.sync();
            const range = ranges.items[0];
            await range.context.sync();
            range.font.highlightColor = "Turquoise";
            const cc = range.insertContentControl();
            cc.tag = "textlint";
            cc.load();
            await cc.context.sync();
            cc.title = (cc.title ? cc.title : "") + result.message + ";";
            await range.context.sync();
          }, []);
        } catch (e) {
          console.log(e);
        }
      }, []);
    } catch (e) {
      console.log(e);
    }

    processing.style.visibility = "hidden";
    proceed.style.visibility = "hidden";

    await context.sync();
  });
}

export async function clear() {
  return Word.run(async context => {
    context.document.body.load("paragraphs");
    await context.sync();
    context.document.body.paragraphs.load("items");
    await context.sync();
    try {
      await context.document.body.paragraphs.items.reduce(async (prev, item) => {
        prev && (await prev);
        item.font.highlightColor = "";
      }, []);

      const ccs = context.document.contentControls.getByTag("textlint");
      ccs.load();
      await context.sync();
      ccs.items.map(cc => cc.delete(true));
    } catch (e) {
      console.log(e);
    }

    await context.sync();
  });
}
