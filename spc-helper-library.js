// ============================================================
// SPC helper library
// Keep helper questions, suggested prompts, and FAQ matching here
// ============================================================

window.SPC_HELPER_LIBRARY = {
  suggestedQuestions: {
    general: [
      "What is SPC?",
      "What is the difference between common and special cause variation?",
      "How do I choose the right chart?",
      "What is a run chart?",
      "What is an XmR chart?",
      "What does stable mean?",
      "What is a target and how should I use it?",
      "How do control limits work?"
    ],
    chart: [
      "What is my chart telling me?",
      "Is my process stable?",
      "Has something changed?",
      "Is it getting better or worse?",
      "What decision should I make?",
      "What about my target?"
    ]
  },

  generalFaq: [
    {
      keywords: ["what is spc", "define spc", ["what", "spc"]],
      answer:
        "Statistical Process Control (SPC) helps you use time-ordered data to understand whether a process is behaving as usual, or whether something has changed.\n\n" +
        "A control chart shows:\n" +
        "• your measure over time\n" +
        "• a centre line (the usual level, often the mean or median)\n" +
        "• control limits (the range you’d expect from routine/common-cause variation)\n\n" +
        "If the pattern breaks simple rules (for example, a point outside the limits or a long run on one side), that’s treated as a special-cause signal — a prompt to investigate what changed in the real world."
    },
    {
      keywords: [
        "common cause",
        "special cause",
        "common and special cause",
        ["difference", "common", "special", "cause"],
        ["what", "common", "cause"],
        ["what", "special", "cause"]
      ],
      answer:
        "Common-cause variation is the routine, expected fluctuation built into the current process.\n\n" +
        "Special-cause variation means something unusual seems to have happened — the pattern suggests the process may have changed.\n\n" +
        "In practice:\n" +
        "• common cause = the usual system working as designed\n" +
        "• special cause = a signal to investigate what changed"
    },
    {
      keywords: [
        "choose the right chart",
        "which chart should i use",
        "how do i choose the right chart",
        ["choose", "chart"]
      ],
      answer:
        "Choose the chart based on the type of data you have:\n\n" +
        "• Run chart: simple time-ordered data, especially when starting out\n" +
        "• XmR chart: continuous data, one value per time point\n" +
        "• C chart: counts of events when the opportunity is roughly constant\n" +
        "• U chart: rates when the opportunity varies\n" +
        "• P chart: proportions\n" +
        "• Xbar-S chart: subgrouped continuous data\n" +
        "• T chart / G chart: rare events\n\n" +
        "A good practical question is: what exactly are you measuring, and what is the denominator or subgroup?"
    },
    {
      keywords: ["run chart", "what is a run chart", ["what", "run", "chart"]],
      answer:
        "A run chart shows data over time with a centre line, usually the median. It helps you see whether the process looks stable or whether there may be a shift or trend.\n\n" +
        "It is often the simplest chart to start with."
    },
    {
      keywords: ["xmr chart", "what is an xmr chart", ["what", "xmr", "chart"]],
      answer:
        "An XmR chart is used for continuous data when you have one observation per time point.\n\n" +
        "It usually has:\n" +
        "• an Individuals chart showing the values over time\n" +
        "• a Moving Range chart showing how much consecutive points differ\n\n" +
        "It helps distinguish routine variation from special-cause signals."
    },
    {
      keywords: [
        "stable",
        "stability",
        "what does stable mean",
        ["what", "stable", "mean"]
      ],
      answer:
        "A stable process shows only routine common-cause variation. In other words, it is predictable within its current level of performance.\n\n" +
        "Stable does not necessarily mean good — it only means the process is behaving consistently."
    },
    {
      keywords: [
        "target",
        "what is a target",
        "how should i use a target",
        ["target", "use"]
      ],
      answer:
        "A target is the performance level you hope or expect to achieve.\n\n" +
        "In SPC, a target is useful for context, but it should not replace understanding whether the process is stable.\n\n" +
        "A stable process below target usually needs redesign or improvement, not just pressure."
    },
    {
      keywords: [
      "median",
      "what is the median",
      ["what", "median"]
     ],
      answer:
    "The median is the middle value when the data are ordered. In run charts it is often used as the centre line because it is simple and robust."
    },
    {
      keywords: [
        "control limits",
        "how do control limits work",
        "ucl",
        "lcl",
        ["control", "limits"],
        ["upper", "lower", "limit"]
      ],
      answer:
        "Control limits show the amount of variation you would expect from routine common-cause variation.\n\n" +
        "They are not the same as targets or specification limits.\n\n" +
        "A point outside the control limits, or a non-random pattern within them, suggests that something may have changed."
    }
  ]
};

// Simple + predictable FAQ matcher
window.matchSpcFaq = function matchSpcFaq(items, text) {
  for (const item of items) {
    const hit = item.keywords.some(k =>
      Array.isArray(k)
        ? k.every(word => text.includes(word))
        : (typeof k === "string" && text.includes(k))
    );

    if (hit) return item.answer;
  }

  return null;
};