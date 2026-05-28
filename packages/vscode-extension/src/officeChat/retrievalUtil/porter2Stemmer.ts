// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const doubleRegex = /(bb|dd|ff|gg|mm|nn|pp|rr|tt)$/;
const nonVowelRegex = /[^aeiouy]/;
const ruleS2: Record<string, string> = {
  ational: "ate",
  ation: "ate",
  ator: "ate",
  tional: "tion",
  enci: "ence",
  anci: "ance",
  izer: "ize",
  abli: "able",
  ization: "ize",
  alism: "al",
  alti: "al",
  alli: "al",
  fulness: "ful",
  ousli: "ous",
  ousness: "ous",
  iviti: "ive",
  iveness: "ive",
  biliti: "ble",
  bli: "ble",
  fulli: "ful",
  lessli: "less",
};

const ruleS3: Record<string, string> = {
  ational: "ate",
  tional: "tion",
  alize: "al",
  icate: "ic",
  iciti: "ic",
  ical: "ic",
  ful: "",
  ness: "",
};

const ruleDeleteS4: string[] = [
  "al",
  "ance",
  "ence",
  "er",
  "ic",
  "able",
  "ible",
  "ant",
  "ement",
  "ment",
  "ent",
  "ism",
  "ate",
  "iti",
  "ous",
  "ive",
  "ize",
];

const ruleSpecialWords: Record<string, string> = {
  skis: "ski",
  skies: "sky",
  dying: "die",
  lying: "lie",
  tying: "tie",
  idly: "idl",
  gently: "gentl",
  ugly: "ugli",
  early: "earli",
  only: "onli",
  singly: "singl",
  sky: "sky",
  news: "news",
  howe: "howe",
  atlas: "atlas",
  cosmos: "cosmos",
  bias: "bias",
  andes: "andes",
};

const exception1a: string[] = [
  "inning",
  "outing",
  "canning",
  "herring",
  "earring",
  "proceed",
  "exceed",
  "succeed",
];

//R1 is the region after the first non-vowel following a vowel, or the end of the word if there is no such non-vowel, mostly.
function getR1(word: string): RegExpMatchArray | null {
  const regException = /^(?:gener|commun|arsen)(.*)/;
  if (regException.test(word)) {
    return word.match(regException);
  }
  const regex = /[aeiouy][^aeiouy](.*)$/;
  const match = word.match(regex);
  return match;
}

function getR2(word: string): RegExpMatchArray | null {
  const r1 = getR1(word);
  if (r1 === null || r1[1].length === 0) {
    return null;
  }
  return getR1(r1[1]);
}

function isShort(word: string): boolean {
  const regShortSyllable1 = /[^aeiouy][aeiouy][^aeiouywxY]/;
  const regShortSyllable2 = /^[aeiouy][^aeiouy]/;
  return (
    regShortSyllable1.test(word) ||
    regShortSyllable2.test(word) ||
    getR1(word) === null ||
    (getR1(word) as RegExpMatchArray)[1].length === 0
  );
}

export function stemmer(value: string): string {
  // check if the word is a special word
  if (value in ruleSpecialWords) {
    return ruleSpecialWords[value];
  }

  //If the word has two letters or less, leave it as it is.
  const word = value.toLowerCase();
  if (word.length < 3) {
    return word;
  }

  //Remove initial '
  while (value.startsWith("'")) {
    value = value.slice(1);
  }

  // Set initial y, or y after a vowel, to Y, and then establish the regions R1 and R2
  if (value.startsWith("y")) {
    value = "Y" + value.slice(1);
  }
  const regY = /([aeiouy])y/g;
  value = value.replace(regY, "$1Y");

  //step 0 Search for the longest among the suffixes, ' 's 's'
  value = value.replace(/'s'$/, "");
  value = value.replace(/s'$/, "");
  value = value.replace(/'$/, "");

  //step 1a Search for the longest among the following suffixes, and perform the action indicated.
  //sses -> ss ied+   ies* replace by i if preceded by more than one letter, otherwise by ie (so ties -> tie, cries -> cri)
  //us+   ss do nothing s delete if the preceding word part contains a vowel not immediately before the s
  const check1a = exception1a.includes(value);
  const regSses = /sses$/;
  const regIes = /(ies|ied)$/;
  const regSfxS = /[aeiouy].+s$/;
  const regSs = /(ss|us)$/;
  if (!check1a) {
    if (regSses.test(value)) {
      value = value.slice(0, -2);
    } else if (regIes.test(value)) {
      value = value.length > 4 ? value.slice(0, -2) : value.slice(0, -1);
    } else if (!regSs.test(value) && regSfxS.test(value)) {
      value = value.slice(0, -1);
    }
  }

  //step 1b Search for the longest among the following suffixes, and, if found, perform the action indicated.
  //eed   eed replace by ee if in R1
  let r1 = getR1(value);
  const regEed = /(eed|eedly)$/;
  const matchELonger1b = regEed.test(value);
  if (r1 && r1[1].length > 0) {
    if (regEed.test(r1[1])) {
      value = value.replace(regEed, "ee");
    }
  }

  //ed    ed  delete if the word contains a vowel, and then
  //if the word ends at, bl or iz add e (so luxuriat -> luxuriate), or
  //if the word ends with a double remove the last letter (so hopp -> hop), or
  //if the word is short, add e (so hop -> hope)
  const regEd = /(ed|edly|ing|ingly)$/;
  if (!matchELonger1b && regEd.test(value)) {
    const preced = value.replace(regEd, "");
    if (nonVowelRegex.test(preced)) {
      value = value.replace(regEd, "");
      if (value.endsWith("at") || value.endsWith("bl") || value.endsWith("iz")) {
        value += "e";
      } else if (doubleRegex.test(value)) {
        const nonAeo = /[^aeo]/;
        if (nonAeo.test(value.slice(0, -2))) {
          value = value.slice(0, -1);
        }
      } else if (isShort(value)) {
        value += "e";
      }
    }
  }

  //step 1c replace suffix y or Y by i if preceded by a non-vowel which is not the first letter of the word (so cry -> cri, by -> by, say -> say)
  const reg1c = /([^aeiouy])[yY]$/;
  if (value.length > 2 && reg1c.test(value)) {
    value = value.slice(0, -1) + "i";
  }

  //step 2 Search for the longest among the following suffixes, and, if found, perform the action indicated.
  r1 = getR1(value);
  if (r1 && r1[1].length > 0) {
    const r1Value = r1[1];
    const regLi = /[cdeghkmnrt]li$/;
    const regLiR1 = /li$/;
    if (regLiR1.test(r1Value) && regLi.test(value)) {
      value = value.slice(0, -2);
    } else if (r1Value.endsWith("ogi") && value.endsWith("logi")) {
      value = value.slice(0, -1);
    } else {
      for (const key in ruleS2) {
        if (r1Value.endsWith(key)) {
          value = value.slice(0, -key.length) + ruleS2[key];
          break;
        }
      }
    }
  }

  //step3 Search for the longest among the following suffixes, and, if found and in R1, perform the action indicated.
  r1 = getR1(value);
  if (r1 && r1[1].length > 0) {
    const r1Value = r1[1];
    const r2 = getR2(value);
    if (r2 && r2[1].length > 0 && r2[1].endsWith("ative")) {
      value = value.slice(0, -5);
    } else {
      for (const key in ruleS3) {
        if (r1Value.endsWith(key)) {
          value = value.slice(0, -key.length) + ruleS3[key];
          break;
        }
      }
    }
  }

  //step4 Search for the longest among the following suffixes, and, if found and in R2, perform the action indicated.
  let r2 = getR2(value);
  if (r2 && r2[1].length > 0) {
    const r2Value = r2[1];
    if (r2Value.endsWith("ion") && (value.endsWith("sion") || r2Value.endsWith("tion"))) {
      value = value.slice(0, -3);
    } else {
      for (const suffix of ruleDeleteS4) {
        if (r2Value.endsWith(suffix)) {
          value = value.slice(0, -suffix.length);
          break;
        }
      }
    }
  }

  //step 5 delete e if in R2, or in R1 and not preceded by a short syllable. delete l if in R2 and preceded by l
  r1 = getR1(value);
  r2 = getR2(value);
  if (r2 && r2[1].length > 0) {
    if (r2[1].endsWith("e")) {
      value = value.slice(0, -1);
    }
  } else if (r1 && r1[1].length > 0) {
    const r1Value = r1[1];
    const regShortSyllable1 = /[^aeiouy][aeiouy]e$/;
    if (r1Value.endsWith("e") && !regShortSyllable1.test(r1Value)) {
      value = value.slice(0, -1);
    }
  }

  r2 = getR2(value);
  if (r2 && r2[1].length > 0) {
    if (r2[1].endsWith("l") && value.endsWith("ll")) {
      value = value.slice(0, -1);
    }
  }

  value = value.replace(/Y/g, "y");

  return value;
}
