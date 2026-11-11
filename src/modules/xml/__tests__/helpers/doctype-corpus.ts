/**
 * DOCTYPE declarations, well-formed and not, each with the verdict of expat (Python's
 * `xml.parsers.expat`) and libxml2 (`xmllint --noout --nonet`) — the two agree on every one. The
 * SAX parser, reading entities (`internalEntities`), must reach the same verdict.
 *
 * `[name, document, wellFormed]`.
 */
export const DOCTYPE_CORPUS: ReadonlyArray<readonly [string, string, boolean]> = [
  ["comment-bracket", '<!DOCTYPE R [ <!-- ] --> ]><R a="1"/>', true],
  ["comment-bracket-gt-tag", '<!DOCTYPE R [ <!-- ] > <F b="x"/> --> ]><R a="1"/>', true],
  ["pi-bracket-gt-tag", '<!DOCTYPE R [ <?pi ] > <F b="x"/> ?> ]><R a="1"/>', true],
  ["comment-apostrophe", "<!DOCTYPE R [ <!-- it's --> ]><R/>", true],
  ["comment-quote", '<!DOCTYPE R [ <!-- say "hi --> ]><R/>', true],
  ["pi-apostrophe", "<!DOCTYPE R [ <?pi it's ?> ]><R/>", true],
  ["entity-literal-bracket", "<!DOCTYPE R [ <!ENTITY e \"]>\"> <!ENTITY f ']>'> ]><R/>", true],
  ["entity-apostrophe", '<!DOCTYPE R [ <!ENTITY e "it\'s"> ]><R/>', true],
  ["attlist-default", '<!DOCTYPE R [ <!ATTLIST R a CDATA "]"> ]><R/>', true],
  [
    "element-decl",
    "<!DOCTYPE R [ <!ELEMENT R (a|b)*> <!ELEMENT a EMPTY> <!ELEMENT b ANY> ]><R/>",
    true
  ],
  ["notation", '<!DOCTYPE R [ <!NOTATION n SYSTEM "x]>"> ]><R/>', true],
  ["notation-public", '<!DOCTYPE R [ <!NOTATION n PUBLIC "-//X//EN"> ]><R/>', true],
  ["system-id", '<!DOCTYPE R SYSTEM "r.dtd"><R/>', true],
  ["public-id", '<!DOCTYPE R PUBLIC "-//X//DTD R//EN" "r.dtd"><R/>', true],
  ["system-and-subset", '<!DOCTYPE R SYSTEM "r.dtd" [ <!ELEMENT R ANY> ]><R/>', true],
  ["space-before-gt", "<!DOCTYPE R [ <!-- ] --> ] ><R/>", true],
  ["pe-decl-and-ref", '<!DOCTYPE R [ <!ENTITY % pe "<!ELEMENT R ANY>"> %pe; ]><R/>', true],
  ["external-entity", '<!DOCTYPE R [ <!ENTITY ext SYSTEM "x.xml"> ]><R/>', true],
  [
    "unparsed-entity",
    '<!DOCTYPE R [ <!NOTATION gif SYSTEM "g"> <!ENTITY pic SYSTEM "p.gif" NDATA gif> ]><R/>',
    true
  ],
  ["entity-used", '<!DOCTYPE R [ <!ENTITY e "x&amp;y"> ]><R a="&e;">&e;</R>', true],
  ["entity-nested", '<!DOCTYPE R [ <!ENTITY a "A"> <!ENTITY b "[&a;]"> ]><R>&b;</R>', true],
  ["entity-charref", '<!DOCTYPE R [ <!ENTITY e "&#38;amp;"> ]><R>&e;</R>', true],
  ["empty-subset", "<!DOCTYPE R []><R/>", true],
  [
    "attlist-default-entity-declared-before",
    '<!DOCTYPE R [ <!ENTITY e "x"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    true
  ],
  [
    "attlist-default-undeclared-with-external-subset",
    '<!DOCTYPE R SYSTEM "r.dtd" [ <!ATTLIST R a CDATA "&missing;"> ]><R/>',
    true
  ],
  [
    "bad-attlist-default-undeclared-entity",
    '<!DOCTYPE R [ <!ATTLIST R a CDATA "&missing;"> ]><R/>',
    false
  ],
  [
    "bad-attlist-default-entity-declared-after",
    '<!DOCTYPE R [ <!ATTLIST R a CDATA "&e;"> <!ENTITY e "x"> ]><R/>',
    false
  ],
  [
    "bad-attlist-default-external-entity",
    '<!DOCTYPE R [ <!ENTITY e SYSTEM "x"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    false
  ],
  ["bad-garbage-in-subset", "<!DOCTYPE R [ garbage ]><R/>", false],
  ["bad-unclosed-comment", "<!DOCTYPE R [ <!-- x ]><R/>", false],
  ["bad-double-dash-comment", "<!DOCTYPE R [ <!-- a -- b --> ]><R/>", false],
  ["bad-unknown-decl", "<!DOCTYPE R [ <!FOO x> ]><R/>", false],
  ["bad-entity-no-value", "<!DOCTYPE R [ <!ENTITY e> ]><R/>", false],
  ["bad-entity-junk", '<!DOCTYPE R [ <!ENTITY e "x" junk> ]><R/>', false],
  ["bad-entity-pe-in-value", '<!DOCTYPE R [ <!ENTITY % p "x"> <!ENTITY e "%p;"> ]><R/>', false],
  [
    "bad-pe-in-decl",
    '<!DOCTYPE R [ <!ENTITY % p "CDATA"> <!ATTLIST R a %p; #IMPLIED> ]><R/>',
    false
  ],
  ["bad-after-subset", "<!DOCTYPE R [ ] x><R/>", false],
  ["bad-no-name", "<!DOCTYPE [ ]><R/>", false],
  ["bad-system-no-literal", "<!DOCTYPE R SYSTEM><R/>", false],
  ["bad-conditional", "<!DOCTYPE R [ <![INCLUDE[ <!ELEMENT R ANY> ]]> ]><R/>", false],
  ["bad-pi-xml-in-subset", '<!DOCTYPE R [ <?xml version="1.0"?> ]><R/>', false],
  ["bad-second-doctype", "<!DOCTYPE R><!DOCTYPE R><R/>", false],
  ["bad-doctype-after-root", "<R/><!DOCTYPE R>", false],
  ["bad-entity-bad-charref", '<!DOCTYPE R [ <!ENTITY e "&#0;"> ]><R/>', false],
  ["bad-entity-recursive", '<!DOCTYPE R [ <!ENTITY a "&b;"> <!ENTITY b "&a;"> ]><R>&a;</R>', false],
  ["bad-entity-undeclared", "<!DOCTYPE R [ ]><R>&nope;</R>", false],
  ["bad-entity-lt-in-attr", '<!DOCTYPE R [ <!ENTITY e "a&#60;b"> ]><R a="&e;"/>', false],
  ["bad-unterminated-literal", '<!DOCTYPE R [ <!ENTITY e "x> ]><R/>', false],
  ["bad-pubid-chars", '<!DOCTYPE R PUBLIC "bad{id}" "r.dtd"><R/>', false],
  ["element-empty", "<!DOCTYPE R [ <!ELEMENT R EMPTY> ]><R/>", true],
  ["element-any", "<!DOCTYPE R [ <!ELEMENT R ANY> ]><R/>", true],
  ["element-mixed-pcdata", "<!DOCTYPE R [ <!ELEMENT R (#PCDATA)> ]><R/>", true],
  ["element-mixed-pcdata-star", "<!DOCTYPE R [ <!ELEMENT R (#PCDATA)*> ]><R/>", true],
  ["element-mixed-names", "<!DOCTYPE R [ <!ELEMENT R ( #PCDATA | a | b )*> ]><R/>", true],
  ["element-seq", "<!DOCTYPE R [ <!ELEMENT R (a, b?, (c | d)+, e*)> ]><R/>", true],
  ["element-nested-groups", "<!DOCTYPE R [ <!ELEMENT R ((a,b)|(c,(d|e)*))?> ]><R/>", true],
  [
    "attlist-many",
    "<!DOCTYPE R [ <!ATTLIST R a CDATA #IMPLIED b ID #REQUIRED c (x|y|1z) 'x' d NOTATION (n) #IMPLIED e NMTOKENS #FIXED \"a b\"> ]><R b='i'/>",
    true
  ],
  ["attlist-no-defs", "<!DOCTYPE R [ <!ATTLIST R> ]><R/>", true],
  ["attlist-default-ref", '<!DOCTYPE R [ <!ATTLIST R a CDATA "x &amp; &#65;"> ]><R/>', true],
  ["notation-public-only", '<!DOCTYPE R [ <!NOTATION n PUBLIC "-//X//EN"> ]><R/>', true],
  ["notation-public-system", '<!DOCTYPE R [ <!NOTATION n PUBLIC "-//X//EN" "x"> ]><R/>', true],
  ["bad-element-mixed-seps", "<!DOCTYPE R [ <!ELEMENT R (a,|b)> ]><R/>", false],
  ["bad-element-mix-separators", "<!DOCTYPE R [ <!ELEMENT R (a,b|c)> ]><R/>", false],
  ["bad-element-empty-group", "<!DOCTYPE R [ <!ELEMENT R ()> ]><R/>", false],
  ["bad-element-no-model", "<!DOCTYPE R [ <!ELEMENT R> ]><R/>", false],
  ["bad-element-unknown-keyword", "<!DOCTYPE R [ <!ELEMENT R NOTHING> ]><R/>", false],
  ["bad-element-pcdata-names-no-star", "<!DOCTYPE R [ <!ELEMENT R (#PCDATA|a)> ]><R/>", false],
  ["bad-element-pcdata-not-first", "<!DOCTYPE R [ <!ELEMENT R (a|#PCDATA)*> ]><R/>", false],
  ["bad-element-unclosed-group", "<!DOCTYPE R [ <!ELEMENT R (a,b> ]><R/>", false],
  ["bad-element-double-occurrence", "<!DOCTYPE R [ <!ELEMENT R (a)**> ]><R/>", false],
  ["bad-attlist-no-type", "<!DOCTYPE R [ <!ATTLIST R a #IMPLIED> ]><R/>", false],
  ["bad-attlist-no-default", "<!DOCTYPE R [ <!ATTLIST R a CDATA> ]><R/>", false],
  ["bad-attlist-bad-type", "<!DOCTYPE R [ <!ATTLIST R a STRING #IMPLIED> ]><R/>", false],
  ["bad-attlist-empty-enum", "<!DOCTYPE R [ <!ATTLIST R a () #IMPLIED> ]><R/>", false],
  ["bad-attlist-lt-in-default", '<!DOCTYPE R [ <!ATTLIST R a CDATA "a<b"> ]><R/>', false],
  ["bad-attlist-fixed-no-value", "<!DOCTYPE R [ <!ATTLIST R a CDATA #FIXED> ]><R/>", false],
  ["bad-attlist-missing-space", '<!DOCTYPE R [ <!ATTLIST R a CDATA"x"> ]><R/>', false],
  ["bad-attlist-notation-no-list", "<!DOCTYPE R [ <!ATTLIST R a NOTATION #IMPLIED> ]><R/>", false],
  ["bad-notation-no-id", "<!DOCTYPE R [ <!NOTATION n> ]><R/>", false],
  ["bad-notation-system-no-literal", "<!DOCTYPE R [ <!NOTATION n SYSTEM> ]><R/>", false],
  ["entity-cdata-end-escaped", '<!DOCTYPE R [ <!ENTITY e "]]&gt;"> ]><R>&e;</R>', true],
  [
    "entity-cdata-end-split",
    '<!DOCTYPE R [ <!ENTITY a "]"> <!ENTITY b "]>"> ]><R>&a;&b;</R>',
    true
  ],
  ["entity-cdata-end-in-attr", '<!DOCTYPE R [ <!ENTITY e "]]>"> ]><R a="&e;"/>', true],
  ["bad-entity-cdata-end", '<!DOCTYPE R [ <!ENTITY e "]]>"> ]><R>&e;</R>', false],
  [
    "bad-entity-cdata-end-charrefs",
    '<!DOCTYPE R [ <!ENTITY e "&#93;&#93;&#62;"> ]><R>&e;</R>',
    false
  ],
  [
    "attlist-default-indirect-escaped-lt",
    '<!DOCTYPE R [ <!ENTITY e "&#38;lt;"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    true
  ],
  [
    "attlist-default-first-binding",
    '<!DOCTYPE R [ <!ENTITY e "ok"> <!ENTITY e SYSTEM "x"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    true
  ],
  [
    "bad-attlist-default-indirect-lt",
    '<!DOCTYPE R [ <!ENTITY e "&#60;"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    false
  ],
  [
    "bad-attlist-default-indirect-undeclared",
    '<!DOCTYPE R [ <!ENTITY e "&missing;"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    false
  ],
  [
    "bad-attlist-default-indirect-recursive",
    '<!DOCTYPE R [ <!ENTITY e "&e;"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    false
  ],
  [
    "bad-attlist-default-indirect-external",
    '<!DOCTYPE R [ <!ENTITY x SYSTEM "x"> <!ENTITY e "&x;"> <!ATTLIST R a CDATA "&e;"> ]><R/>',
    false
  ],
  [
    "standalone-declarations-after-pe",
    '<?xml version="1.0" standalone="yes"?><!DOCTYPE R [ <!ENTITY % p SYSTEM "x"> %p; <!ENTITY e "E"> ]><R>&e;</R>',
    true
  ],
  [
    "bad-standalone-attlist-default-undeclared",
    '<?xml version="1.0" standalone="yes"?><!DOCTYPE R SYSTEM "x" [ <!ATTLIST R a CDATA "&missing;"> ]><R/>',
    false
  ]
];
