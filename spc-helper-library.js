// ============================================================
// SPC helper library
// Designed for healthcare service improvement teams
// Plain English, but SPC-defensible
// ============================================================

(function () {
  function norm(text) {
    return String(text || "")
      .toLowerCase()
      .replace(/[’']/g, "'")
      .replace(/[^a-z0-9%+\-.\s]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function includesAll(text, parts) {
    return parts.every(p => text.includes(norm(p)));
  }

  function includesAny(text, parts) {
    return parts.some(p => text.includes(norm(p)));
  }

  function finiteNumber(x) {
    return typeof x === "number" && Number.isFinite(x);
  }

  function titleCaseChartType(chartType) {
    switch (chartType) {
      case "run": return "Run chart";
      case "xmr": return "XmR chart";
      case "xbars": return "X̄–S chart";
      case "c": return "C chart";
      case "p": return "P chart";
      case "u": return "U chart";
      case "t": return "T chart";
      case "g": return "G chart";
      default: return "chart";
    }
  }

  function signalTextList(signals) {
    if (!Array.isArray(signals) || signals.length === 0) return [];

    return signals.map(s => {
      const t = String(s || "").toLowerCase();

      if (t.includes("above ucl")) return "one or more points above the upper limit";
      if (t.includes("below lcl")) return "one or more points below the lower limit";
      if (t.includes("run")) return "a sustained run on one side of the centre line";
      if (t.includes("trend")) return "a sustained trend";
      if (t.includes("astronomical")) return "an unusual outlier";
      if (t.includes("two of three")) return "an unusual clustering of points far from the centre line";
      if (t.includes("four of five")) return "a sustained bias away from the centre line";

      return s;
    });
  }

  function joinNice(items) {
    const arr = (items || []).filter(Boolean);
    if (arr.length === 0) return "";
    if (arr.length === 1) return arr[0];
    if (arr.length === 2) return `${arr[0]} and ${arr[1]}`;
    return `${arr.slice(0, -1).join(", ")}, and ${arr[arr.length - 1]}`;
  }

  function getChartTypeSafe() {
    try {
      if (typeof getSelectedChartType_NoSideEffects === "function") {
        return getSelectedChartType_NoSideEffects() || "run";
      }
      if (typeof getSelectedChartType === "function") {
        return getSelectedChartType() || "run";
      }
    } catch (e) {}
    return "run";
  }

  function getCurrentChartContext() {
    const chartType = getChartTypeSafe();

    return {
      chartType,
      run: typeof lastRunAnalysis !== "undefined" ? lastRunAnalysis : null,
      xmr: typeof lastXmRAnalysis !== "undefined" ? lastXmRAnalysis : null,
      xbars: typeof lastXbarSAnalysis !== "undefined" ? lastXbarSAnalysis : null,
      attribute: typeof lastAttributeAnalysis !== "undefined" ? lastAttributeAnalysis : null,
      rare: typeof lastRareAnalysis !== "undefined" ? lastRareAnalysis : null,
      xmrPeriods: window.lastXmRPeriods || null
    };
  }

  function getSuggestedQuestions(hasChart) {
    return {
      general: [
        "What is SPC?",
        "What is the difference between common and special cause variation?",
        "How do I choose the right chart?",
        "What is a run chart?",
        "What is an XmR chart?",
        "What does stable mean?",
        "What is the difference between a baseline and a split?",
        "How do control limits work?",
        "What should I do if I see a signal?",
        "What if my process is stable but performance is poor?"
      ],
      chart: [
        "What is my chart telling me?",
        "Has anything changed?",
        "Is my process stable?",
        "Is this likely to be real improvement?",
        "What does this mean for our target?",
        "What should I do next?",
        "Should I add a split here?"
      ].filter(Boolean)
    };
  }

  const FAQ = [
    {
      id: "what-is-spc",
      priority: 8,
      aliases: ["what is spc", "define spc", "explain spc"],
      keywords: [["what", "spc"], ["define", "spc"], ["explain", "spc"]],
      answer:
        "Statistical Process Control (SPC) helps you use data over time to understand whether a process is behaving as usual, or whether something may have changed.\n\n" +
        "In healthcare improvement, that helps teams distinguish routine ups and downs from signals that may reflect a real change in the service, pathway, staffing, demand, coding, or measurement system."
    },
    {
      id: "common-vs-special",
      priority: 8,
      aliases: ["common cause", "special cause", "difference between common and special cause variation"],
      keywords: [
        ["common", "cause"],
        ["special", "cause"],
        ["difference", "common", "special", "cause"]
      ],
      answer:
        "Common-cause variation is the routine fluctuation built into the current system. Special-cause variation means the pattern suggests something unusual may have happened.\n\n" +
        "In practice:\n" +
        "• common cause = the system behaving as it usually does\n" +
        "• special cause = a prompt to investigate what may have changed"
    },
    {
      id: "choose-chart",
      priority: 8,
      aliases: ["choose the right chart", "which chart should i use", "how do i choose the right chart"],
      keywords: [
        ["choose", "chart"],
        ["which", "chart"],
        ["right", "chart"]
      ],
      answer:
        "Choose the chart based on the type of data you have:\n\n" +
        "• Run chart: simple data over time, often a good starting point\n" +
        "• XmR chart: continuous data with one value per time point\n" +
        "• C chart: counts with roughly constant opportunity\n" +
        "• P chart: proportions\n" +
        "• U chart: rates with changing opportunity\n" +
        "• X̄–S chart: subgrouped continuous data\n" +
        "• T / G charts: rare events\n\n" +
        "In healthcare, it often helps to ask: what exactly is the numerator, denominator, subgroup, or opportunity?"
    },
    {
      id: "run-chart",
      priority: 7,
      aliases: ["what is a run chart", "run chart"],
      keywords: [["run", "chart"]],
      answer:
        "A run chart shows data over time with a centre line, usually the median. It is a simple way to see whether the pattern looks broadly stable or whether there may be a shift or trend.\n\n" +
        "It is often a very good starting point for healthcare improvement teams."
    },
    {
      id: "xmr-chart",
      priority: 7,
      aliases: ["what is an xmr chart", "xmr chart"],
      keywords: [["xmr", "chart"]],
      answer:
        "An XmR chart is used for continuous data when you have one observation per time point.\n\n" +
        "It usually has:\n" +
        "• an Individuals (X) chart showing the values over time\n" +
        "• a Moving Range (MR) chart showing how much consecutive points differ\n\n" +
        "It helps separate routine variation from possible special-cause signals."
    },
    {
      id: "stable-meaning",
      priority: 7,
      aliases: ["what does stable mean", "what is a stable process", "stability"],
      keywords: [["stable"], ["stability"]],
      answer:
        "A stable process shows only routine common-cause variation. In simple terms, it is behaving predictably within its current level of performance.\n\n" +
        "Stable does not necessarily mean good. A stable process can still be consistently underperforming."
    },
    {
      id: "control-limits",
      priority: 8,
      aliases: ["control limits", "ucl", "lcl", "how do control limits work"],
      keywords: [["control", "limit"], ["ucl"], ["lcl"], ["upper", "lower", "limit"]],
      answer:
        "Control limits show the range you would expect from routine common-cause variation. They are not the same as targets, standards, or pass/fail thresholds.\n\n" +
        "A point outside the limits, or a clear non-random pattern within them, suggests something may have changed and is worth investigating."
    },
    {
      id: "target",
      priority: 6,
      aliases: ["what is a target", "how should i use a target", "target"],
      keywords: [["target"]],
      answer:
        "A target is the level of performance you are aiming for. In SPC, it should be interpreted alongside stability.\n\n" +
        "A stable process below target often means the system needs redesign or improvement. An unstable process meeting target intermittently may still not be reliable."
    },
    {
      id: "baseline",
      priority: 7,
      aliases: ["what is a baseline", "when should i use a baseline"],
      keywords: [["baseline"]],
      answer:
        "A baseline is the initial set of points used to calculate the centre line and, where relevant, the limits.\n\n" +
        "It is useful when you want to compare later performance against an earlier period that represents the original system."
    },
    {
      id: "baseline-vs-split",
      priority: 8,
      aliases: ["difference between a baseline and a split", "baseline and split", "baseline vs split"],
      keywords: [
        ["baseline", "split"],
        ["difference", "baseline", "split"]
      ],
      answer:
        "A baseline fixes the calculation using an earlier reference period. A split recalculates the centre line and limits from a chosen point onward.\n\n" +
        "In simple terms:\n" +
        "• baseline = compare later performance against the old system\n" +
        "• split = treat the later period as a new system or new normal"
    },
    {
      id: "what-to-do-signal",
      priority: 8,
      aliases: ["what should i do if i see a signal", "what do i do if i see a signal", "what should i do next"],
      keywords: [
        ["what", "do", "signal"],
        ["what", "should", "i", "do", "next"],
        ["see", "signal"]
      ],
      answer:
        "A special-cause signal is usually a prompt to investigate what changed around that time.\n\n" +
        "In a healthcare service, that might include a pathway change, staffing change, demand pressure, case-mix change, coding change, data-definition change, equipment issue, or an intentional improvement intervention.\n\n" +
        "The aim is not to blame a person or overreact to one point. The aim is to understand whether something in the system changed, whether that change is real, and whether it is likely to continue."
    },
    {
      id: "stable-but-poor",
      priority: 8,
      aliases: ["stable but poor", "stable but below target", "stable but not good enough"],
      keywords: [
        ["stable", "poor"],
        ["stable", "below", "target"],
        ["stable", "not", "good", "enough"]
      ],
      answer:
        "If the process is stable but the level of performance is not good enough, that usually means the system is delivering a predictable but unsatisfactory result.\n\n" +
        "In improvement terms, that often points toward redesigning or improving the system rather than reacting to individual good or bad days."
    },
    {
      id: "data-definition-change",
      priority: 7,
      aliases: ["what if our data definition changed", "coding change", "case definition changed"],
      keywords: [
        ["data", "definition", "changed"],
        ["coding", "change"],
        ["case", "definition", "changed"]
      ],
      answer:
        "A change in coding, inclusion criteria, denominator rules, or case definition can create a chart signal even if the real service has not changed.\n\n" +
        "That is why SPC charts should always be interpreted alongside local knowledge about data collection and operational context."
    },
    {
      id: "annotations",
      priority: 6,
      aliases: ["how should i use annotations", "annotations"],
      keywords: [["annotation"], ["annotations"]],
      answer:
        "Use annotations to mark events that could plausibly explain a change in the pattern, such as a new pathway, staffing change, policy change, or data-definition change.\n\n" +
        "Keep them short and factual. They support interpretation, but they do not prove causation on their own."
    },
    {
      id: "improvement-vs-one-good-point",
      priority: 7,
      aliases: ["are we improving", "is this improvement", "one good point"],
      keywords: [
        ["improv"],
        ["getting", "better"],
        ["one", "good", "point"]
      ],
      answer:
        "In SPC, improvement is more convincing when the chart shows a sustained change in the desired direction, not just one good point.\n\n" +
        "A single good result can happen by chance. A sustained shift or other clear signal is more consistent with real change."
    }
  ];

  function scoreFaqItem(item, q) {
    let score = item.priority || 0;

    (item.aliases || []).forEach(alias => {
      const a = norm(alias);
      if (q === a) score += 100;
      else if (q.includes(a)) score += 30;
    });

    (item.keywords || []).forEach(k => {
      if (Array.isArray(k)) {
        if (includesAll(q, k)) score += 15;
      } else {
        const kk = norm(k);
        if (q.includes(kk)) score += 10;
      }
    });

    return score;
  }

  function findBestFaq(q) {
    let best = null;
    let bestScore = 0;

    FAQ.forEach(item => {
      const score = scoreFaqItem(item, q);
      if (score > bestScore) {
        best = item;
        bestScore = score;
      }
    });

    return bestScore >= 15 ? best : null;
  }

  function describeXmRPeriods(periods) {
    if (!Array.isArray(periods) || periods.length === 0) return null;

    const pieces = periods.map((p, idx) => {
      const label = `Period ${idx + 1}`;
      const start = (p.startIndex ?? 0) + 1;
      const end = (p.endIndex ?? 0) + 1;
      const stable = !!p.isStable;
      const sigs = signalTextList(p.signals);
      const meanTxt = finiteNumber(p.mean) ? p.mean.toFixed(2) : null;

      if (stable) {
        return `${label} (points ${start}–${end}) looks stable${meanTxt ? ` around a mean of ${meanTxt}` : ""}`;
      }

      return `${label} (points ${start}–${end}) shows ${joinNice(sigs) || "special-cause signals"}${meanTxt ? ` around a mean of ${meanTxt}` : ""}`;
    });

    return pieces.join(". ") + ".";
  }

  function buildXmRChangeAnswer(ctx) {
    const latest = ctx.xmr;
    const periods = ctx.xmrPeriods;

    if (Array.isArray(periods) && periods.length > 1) {
      const earlierUnstable = periods.slice(0, -1).some(p => !p.isStable);
      const latestStable = !!periods[periods.length - 1].isStable;
      const latestStart = (periods[periods.length - 1].startIndex ?? 0) + 1;
      const latestEnd = (periods[periods.length - 1].endIndex ?? 0) + 1;

      let out = "";

      if (earlierUnstable) {
        out += "Yes — looking across the whole XmR chart, there is evidence that performance changed at some point, because one or more earlier periods show special-cause signals or a re-based pattern after a split.\n\n";
      } else {
        out += "Looking across the whole XmR chart, there is no strong evidence of a major change between periods, although the chart may still have been intentionally split for interpretation.\n\n";
      }

      out += describeXmRPeriods(periods) + "\n\n";

      if (latestStable) {
        out += `In the most recent period (points ${latestStart}–${latestEnd}), I cannot see a clear signal that the process has changed again. That latest section looks like routine variation within its current level.`;
      } else {
        const latestSignals = joinNice(signalTextList(periods[periods.length - 1].signals));
        out += `In the most recent period (points ${latestStart}–${latestEnd}), there are still signs of possible change — specifically ${latestSignals || "special-cause signals"}.`;
      }

      return out;
    }

    if (!latest) {
      return "Please generate an XmR chart first, then ask again.";
    }

    if (latest.isStable) {
      return "Looking at the current XmR chart, I cannot see a clear special-cause signal. The process looks broadly stable in its current form.";
    }

    return "Yes — this XmR chart shows evidence that something may have changed. The pattern is not fully explained by routine variation alone, so it would be worth investigating what changed in the system or the data around that time.";
  }

  function buildLatestPeriodSummary(ctx) {
    const chartType = ctx.chartType;

    if (chartType === "xmr" && ctx.xmr) {
      const a = ctx.xmr;
      const start = (a.startIndex ?? 0) + 1;
      const end = (a.endIndex ?? 0) + 1;
      const periodText = (a.periodCount && a.periodCount > 1)
        ? `Looking at the latest period (Period ${a.periodIndex} of ${a.periodCount}, points ${start}–${end})`
        : `Looking at the chart (points ${start}–${end})`;

      if (a.isStable) {
        return `${periodText}, this XmR chart looks stable overall. I cannot see a clear special-cause signal using the current rule settings.`;
      }

      const sigs = joinNice(signalTextList(a.signals));
      return `${periodText}, this XmR chart shows possible special-cause variation — specifically ${sigs || "one or more SPC signals"}.`;
    }

    if (chartType === "xbars" && ctx.xbars) {
      const xStable = !!ctx.xbars.xbar?.isStable;
      const sStable = !!ctx.xbars.s?.isStable;

      if (xStable && sStable) {
        return "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation look stable.";
      }
      if (!xStable && sStable) {
        return "Looking at the latest X̄–S period, the subgroup averages show possible change, but the within-subgroup variation looks stable.";
      }
      if (xStable && !sStable) {
        return "Looking at the latest X̄–S period, the subgroup averages look stable, but the within-subgroup variation shows possible change.";
      }
      return "Looking at the latest X̄–S period, both the subgroup averages and the within-subgroup variation show possible change.";
    }

    if (ctx.attribute) {
      if (ctx.attribute.isStable) {
        return `Looking at the latest ${titleCaseChartType(chartType).toLowerCase()} period, I cannot see a clear special-cause signal.`;
      }
      return `Looking at the latest ${titleCaseChartType(chartType).toLowerCase()} period, there are signs of possible special-cause variation.`;
    }

    if (ctx.run) {
      if (ctx.run.isStable) {
        return "Looking at the current run chart, I cannot see a clear signal that the pattern has changed.";
      }
      const sigs = joinNice(signalTextList(ctx.run.signals));
      return `Looking at the current run chart, there are signs of possible change — specifically ${sigs || "a non-random pattern"}.`;
    }

    return "Generate a chart first and I can give you a chart-specific interpretation.";
  }

  function buildTargetAnswer(ctx) {
    const chartType = ctx.chartType;

    if (chartType === "xmr" && ctx.xmr) {
      if (ctx.xmr.isStable) {
        return "In the latest XmR period, the process looks stable, so the target becomes more useful for judging whether the current system is reliably good enough.\n\nIf the stable process is still below target, that usually suggests the system needs improvement or redesign rather than pressure on individual teams or shifts.";
      }
      return "In the latest XmR period, the process does not look fully stable, so target performance should be interpreted cautiously.\n\nThe first question is often what is changing in the system, demand, staffing, coding, or measurement — not just whether the target was hit on isolated points.";
    }

    return "Targets are most useful when interpreted alongside stability. A stable process below target often needs system redesign. An unstable process meeting target occasionally may still not be reliable.";
  }

  function buildWhatChartTellingMe(ctx) {
    const chartType = ctx.chartType;

    if (chartType === "xmr" && ctx.xmr) {
      const overall = buildXmRChangeAnswer(ctx);
      return overall;
    }

    return buildLatestPeriodSummary(ctx);
  }

  function buildDecisionAnswer(ctx) {
    const latestSummary = buildLatestPeriodSummary(ctx);

    return latestSummary + "\n\n" +
      "From a healthcare improvement point of view, the next step is usually:\n" +
      "• if the process looks stable but performance is not good enough, think about redesigning the system\n" +
      "• if the chart shows a signal, investigate what changed in the service, pathway, demand, staffing, coding, or measurement\n" +
      "• avoid reacting to individual points as if each one proves success or failure";
  }

  function answerQuestion(question) {
    const qRaw = String(question || "").trim();
    const q = norm(qRaw);

    if (!q) {
      return "Please type a question about SPC or your chart.";
    }

    const ctx = getCurrentChartContext();

    // Prefer chart-aware answers for clearly chart-specific questions
    const asksWhatChartSays =
      includesAll(q, ["what", "chart", "telling"]) ||
      includesAll(q, ["what", "my", "chart"]) ||
      includesAll(q, ["interpret", "chart"]);

    const asksChanged =
      includesAll(q, ["has", "anything", "changed"]) ||
      includesAll(q, ["has", "something", "changed"]) ||
      includesAll(q, ["significant", "change"]) ||
      includesAll(q, ["changed"]);

    const asksStable =
      includesAll(q, ["stable"]) ||
      includesAll(q, ["process", "stable"]) ||
      includesAll(q, ["is", "my", "process", "stable"]);

    const asksTargetMeaning =
      includesAll(q, ["target"]) &&
      (includesAny(q, ["mean", "meaning", "what does", "what about", "tell me"]) || q.includes("target"));

    const asksDecision =
      includesAll(q, ["what", "should", "i", "do"]) ||
      includesAll(q, ["what", "decision"]) ||
      includesAll(q, ["next", "step"]);

    const hasChart =
      !!ctx.run || !!ctx.xmr || !!ctx.xbars || !!ctx.attribute || !!ctx.rare;

    if (hasChart && asksWhatChartSays) return buildWhatChartTellingMe(ctx);
    if (hasChart && asksChanged && ctx.chartType === "xmr") return buildXmRChangeAnswer(ctx);
    if (hasChart && asksStable) return buildLatestPeriodSummary(ctx);
    if (hasChart && asksTargetMeaning) return buildTargetAnswer(ctx);
    if (hasChart && asksDecision) return buildDecisionAnswer(ctx);

    // General FAQ
    const faq = findBestFaq(q);
    if (faq) return faq.answer;

    // Fallback
    if (hasChart) {
      return "I’m not fully sure which answer fits best. Try one of these:\n" +
        "• What is my chart telling me?\n" +
        "• Has anything changed?\n" +
        "• Is my process stable?\n" +
        "• What should I do next?\n" +
        "• What does this mean for our target?";
    }

    return "I’m not fully sure which answer fits best. Try one of these:\n" +
      "• What is SPC?\n" +
      "• How do I choose the right chart?\n" +
      "• What is the difference between a baseline and a split?\n" +
      "• What should I do if I see a signal?\n" +
      "• What if my process is stable but performance is poor?";
  }

  window.SPC_HELPER_LIBRARY = {
    getSuggestedQuestions,
    answerQuestion,
    faq: FAQ
  };

  window.matchSpcFaq = function matchSpcFaq(items, text) {
    const q = norm(text);
    let best = null;
    let bestScore = 0;

    (items || []).forEach(item => {
      const score = scoreFaqItem(item, q);
      if (score > bestScore) {
        best = item;
        bestScore = score;
      }
    });

    return best && bestScore >= 15 ? best.answer : null;
  };
})();