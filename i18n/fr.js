
/*
 * SimpleSPC French translations
 *
 * Initial French dictionary.
 * Terminology will be reviewed with a French-speaking
 * SPC practitioner before the bilingual release.
 */

window.SimpleSPCi18n.registerLanguage("fr", {

  // Application heading
  "app.title": "Outil Web Simple SPC",

  "app.subtitle": "Importez des données et créez des cartes de contrôle statistique directement dans votre navigateur. Vos données ne quittent jamais votre navigateur.",

  // Main buttons
  "buttons.generate": "Générer / Recalculer",
  "buttons.reset": "Réinitialiser le graphique",
  "buttons.help": "Aide et conseils",
  "buttons.spcHelper": "Assistant SPC",

  // Main tabs
  "tabs.data": "Données",
  "tabs.chart": "Graphique",
  "tabs.labels": "Libellés",
  "tabs.axes": "Axes",
  "tabs.baseline": "Référence",
  
  "tabs.rules": "Règles",

  // Sélecteur de langue
  "language.label": "Langue",

  // Introduction
  "app.intro": "Ajoutez vos données, choisissez votre graphique et ses paramètres, puis cliquez sur « Générer / Recalculer » si nécessaire. Vous pouvez importer des fichiers CSV ou Excel, modifier les données dans l'éditeur et enregistrer votre graphique pour le rouvrir plus tard. Vous avez des questions sur la maîtrise statistique des procédés ? Demandez à l'assistant.",

  // Infobulles
  "buttons.resetTitle": "Conserver les données chargées et les colonnes sélectionnées, mais réinitialiser les paramètres du graphique, notamment les objectifs, les changements de phase, les annotations, les règles et la mise en forme.",

  
"tabs.baselineTitle": "Définir une période de référence initiale ou un objectif de performance",

// PANNEAU DONNÉES
"data.addHeading": "Ajouter des données :",
"data.requirementsAria": "Aide sur les données nécessaires",
"data.requirementsTooltip": "Les colonnes nécessaires dépendent du type de graphique. La plupart des cartes SPC nécessitent une colonne de temps ou de séquence et une mesure. Certaines nécessitent également des nombres d'événements et des dénominateurs. Les diagrammes de Pareto nécessitent une colonne de catégories et peuvent également utiliser une colonne d'effectifs.",
"data.uploadHint": "Importez un fichier CSV/Excel ou collez vos données dans l'éditeur.",
"data.openEditor": "Ouvrir l'éditeur de données (coller/saisir)",
"data.clear": "Effacer les données et le graphique",
"data.clearTitle": "Supprimer les données chargées et réinitialiser le graphique. Enregistrez votre travail au préalable si vous souhaitez le rouvrir plus tard.",

"data.saveHeading": "Enregistrer ou ouvrir un graphique",
"data.saveHint": "Enregistrez votre graphique, vos données et vos paramètres dans un fichier JSON pour pouvoir les rouvrir et les modifier ultérieurement.",
"data.saveButton": "Enregistrer le graphique",
"data.openButton": "Ouvrir un graphique enregistré",

"data.portableHeading": "Enregistrer une copie hors ligne",
"data.portableAria": "Aide sur les graphiques portables",
"data.portableTooltip": "Enregistrez une copie HTML autonome de votre graphique. Ce fichier contient vos données et une version fonctionnelle de SimpleSPC. Vous pouvez l'ouvrir sans connexion Internet, apporter des modifications et l'enregistrer à nouveau. Les fichiers portables contiennent vos données. Stockez-les et partagez-les conformément aux exigences de gouvernance de l'information de votre organisation.",
"data.portableButton": "Enregistrer un graphique portable",

// PANNEAU GRAPHIQUE
"chart.noData": "Chargez des données pour activer le choix du graphique et des colonnes.",
"chart.newDataHeading": "Nouvelles données chargées",
"chart.newDataQuestion": "Souhaitez-vous être guidé pour choisir le graphique le plus adapté à ces données ?",
"chart.helpChoose": "Aidez-moi à choisir",
"chart.chooseManually": "Choisir manuellement",
"chart.typeHeading": "Type de graphique :",
"chart.run": "Graphique chronologique",
"chart.xmr": "Carte XmR",
"chart.moreTypes": "Autres types de graphiques",

"chart.setup": "Paramètres du graphique",

// CONTRÔLES DES CARTES SECONDAIRES MR / S
"chart.secondary.showMR": "Afficher la carte MR",
"chart.secondary.showS": "Afficher la carte S",
"chart.secondary.mrDisplay": "Affichage MR",
"chart.secondary.sDisplay": "Affichage de la carte S",
"chart.secondary.includesSplits": "(phases incluses)",
"chart.secondary.current": "Période actuelle",
"chart.secondary.all": "Toutes",
"chart.secondary.helpAria": "Aide sur l'affichage de la carte secondaire",
"chart.secondary.helpText": "Les cartes XmR peuvent afficher une carte des étendues mobiles (MR), qui permet d'étudier la variation entre deux mesures consécutives. Les cartes X̄–S peuvent afficher une carte S, qui représente la variation au sein des sous-groupes. Période actuelle affiche uniquement la dernière période d'analyse. Toutes affiche l'historique complet, y compris les différentes phases. Utilisez la case à cocher pour afficher ou masquer la carte secondaire.",

// SÉLECTION DES COLONNES
"chart.columns.heading": "Choisir les colonnes :",
"chart.columns.xDefault": "Date / colonne de l'axe X",
"chart.columns.yDefault": "Valeur / colonne de l'axe Y",
"chart.columns.thirdDefault": "Dénominateur / opportunités",
"chart.columns.xmrY": "Mesure (utilisée pour les limites XmR)",
"chart.columns.cY": "Nombre (c) par période",
"chart.columns.pY": "Numérateur : unités non conformes (d)",
"chart.columns.pThird": "Dénominateur : total (n)",
"chart.columns.pHint": "La carte P représente une proportion : d sur n.",
"chart.columns.pprimeHint": "La carte P′ représente une proportion : d sur n, avec des limites de contrôle ajustées selon la méthode de Laney.",
"chart.columns.uY": "Numérateur : non-conformités (c)",
"chart.columns.uThird": "Dénominateur : opportunités (n)",
"chart.columns.uHint": "La carte U représente le nombre de non-conformités par opportunité : c sur n.",
"chart.columns.uprimeHint": "La carte U′ représente le nombre de non-conformités par opportunité, avec des limites de contrôle ajustées selon la méthode de Laney.",
"chart.columns.xbarsY": "Valeur mesurée",
"chart.columns.xbarsThird": "Identifiant du sous-groupe (jour / semaine / échantillon)",
"chart.columns.xbarsHint": "La carte X̄–S nécessite plusieurs mesures par sous-groupe. L'axe horizontal utilise automatiquement l'ordre des sous-groupes.",
"chart.columns.tGapsY": "Temps entre les événements (par ex. en jours)",
"chart.columns.tDatesY": "Colonne de valeurs non utilisée (la carte T utilise les dates)",
"chart.columns.tGapsHint": "La carte T représente le temps entre des événements rares à partir des intervalles numériques fournis.",
"chart.columns.tDatesHint": "La carte T représente le temps entre des événements rares, calculé à partir des dates des événements.",
"chart.columns.gY": "Opportunités entre les événements",
"chart.columns.gHint": "La carte G représente le nombre d'opportunités entre des événements rares.",
"chart.columns.paretoX": "Colonne des catégories",

"chart.columns.paretoY": "Colonne des fréquences / effectifs",

// AUTRES TYPES DE GRAPHIQUES
"chart.types.c": "Carte C (dénombrements)",
"chart.types.p": "Carte P (proportions)",
"chart.types.pprime": "Carte P′ (Laney)",
"chart.types.u": "Carte U (taux)",
"chart.types.uprime": "Carte U′ (Laney)",
"chart.types.xbars": "Carte X̄–S (sous-groupes)",
"chart.types.t": "Carte T (temps entre événements)",
"chart.types.g": "Carte G (opportunités entre événements)",
"chart.types.pareto": "Diagramme de Pareto (catégories)",

// AFFICHAGE DES TAUX
"chart.rate.label": "Afficher le taux pour :",
"chart.rate.per1": "Pour 1",
"chart.rate.per10": "Pour 10",
"chart.rate.per100": "Pour 100",
"chart.rate.per1000": "Pour 1 000",
"chart.rate.per10000": "Pour 10 000",
"chart.rate.hint": "Modifie uniquement l'affichage du taux ; le calcul sous-jacent de la carte U reste inchangé.",

// TYPE D'AXE HORIZONTAL
"chart.axis.type": "Type d'axe horizontal",
"chart.axis.helpAria": "Aide sur le type d'axe horizontal",
"chart.axis.help": "Utilisez Date / heure pour les cartes SPC chronologiques. Utilisez Séquence / catégorie pour afficher les observations dans leur ordre initial sans interpréter les valeurs de l'axe horizontal comme des dates.",
"chart.axis.date": "Date / heure",
"chart.axis.sequence": "Séquence / catégorie",

// FORMAT DES DATES
"chart.dateFormat.label": "Format des dates numériques",
"chart.dateFormat.helpAria": "Aide sur le format des dates",
"chart.dateFormat.help": "Utilisé pour les dates numériques ambiguës, comme 01/02/2024. Les dates au format ISO, comme 2024-02-01, et les dates Excel sont reconnues automatiquement.",
"chart.dateFormat.uk": "Royaume-Uni / jour d'abord (jj/mm/aaaa)",
"chart.dateFormat.us": "États-Unis / mois d'abord (mm/jj/aaaa)",
"chart.dateFormat.auto": "Détection automatique",

"chart.dateFormat.iso": "Dates ISO / Excel uniquement",

// DISPONIBILITÉ DES CARTES
"chart.availability.noData": "Chargez des données pour déterminer les types de cartes disponibles.",
"chart.availability.selectedColumn": "la colonne de valeurs sélectionnée",
"chart.availability.xmr": "Moins de 12 points de données sont disponibles dans « {column} ». Un minimum de 12 points est nécessaire pour une carte XmR valide.",
"chart.availability.c": "Aucune colonne appropriée de nombres entiers non négatifs n'est disponible. Les cartes C nécessitent des dénombrements par période.",
"chart.availability.p": "Aucune paire numérateur/dénominateur appropriée n'est disponible. Les cartes P nécessitent deux colonnes de nombres entiers, le numérateur représentant une partie du dénominateur.",
"chart.availability.pprime": "Aucune paire numérateur/dénominateur appropriée n'est disponible. Les cartes P′ nécessitent deux colonnes de nombres entiers, le numérateur représentant une partie du dénominateur.",
"chart.availability.u": "Aucune paire nombre d'événements/opportunités appropriée n'est disponible. Les cartes U nécessitent une colonne de dénombrements entiers et une colonne d'opportunités strictement positives.",
"chart.availability.uprime": "Aucune paire nombre d'événements/opportunités appropriée n'est disponible. Les cartes U′ nécessitent une colonne de dénombrements entiers et une colonne d'opportunités strictement positives.",
"chart.availability.xbars": "Les cartes X̄–S nécessitent une colonne de mesures numériques et une colonne de sous-groupes contenant des identifiants répétés.",
"chart.availability.t": "Les cartes T nécessitent des dates d'événements ou une colonne numérique indiquant le temps entre les événements.",
"chart.availability.g": "Les cartes G nécessitent une colonne de nombres entiers dont toutes les valeurs sont au moins égales à 1, car elles représentent le nombre d'opportunités entre des événements rares.",

// AVERTISSEMENTS SUR LES DATES
"chart.warning.mixedDates": "Attention : cette colonne semble contenir un mélange de dates numériques au format britannique et américain. Harmonisez les formats si possible.",
"chart.warning.usDatesUkSelected": "Attention : ces dates semblent utiliser le format américain, mois/jour/année (mm/jj/aaaa), alors que le format britannique, jour/mois/année, est sélectionné. Certaines lignes pourraient être ignorées. Essayez le format américain ou la détection automatique.",
"chart.warning.ukDatesUsSelected": "Attention : ces dates semblent utiliser le format britannique, jour/mois/année (jj/mm/aaaa), alors que le format américain, mois/jour/année, est sélectionné. Certaines lignes pourraient être ignorées. Essayez le format britannique ou la détection automatique.",
"chart.warning.ambiguousUk": "Cette colonne contient des dates numériques ambiguës. Elles seront actuellement interprétées au format britannique, jour/mois/année (jj/mm/aaaa).",
"chart.warning.ambiguousUs": "Cette colonne contient des dates numériques ambiguës. Elles seront actuellement interprétées au format américain, mois/jour/année (mm/jj/aaaa).",
"chart.warning.ambiguousAuto": "Cette colonne contient des dates numériques ambiguës et l'application ne peut pas déterminer leur format avec suffisamment de certitude. Elle utilisera le format britannique par défaut, sauf si vous choisissez une autre option.",
"chart.warning.ambiguousIso": "Cette colonne contient des dates numériques avec des barres obliques ou des traits d'union. En mode « Dates ISO / Excel uniquement », les dates numériques ambiguës peuvent ne pas être reconnues comme des dates.",

// CONTRÔLE DU FORMAT DES DATES DÉSACTIVÉ

"chart.dateFormat.disabled": "Le choix du format des dates est disponible uniquement lorsque l'axe horizontal utilise Date / heure.",

// LIBELLÉS COMMUNS DE MISE EN FORME
"common.autoDefault": "Auto / par défaut",
"common.auto": "Auto",
"common.bold": "Gras",
"common.italic": "Italique",
"common.underline": "Souligné",

// PANNEAU LIBELLÉS
"labels.heading": "Libellés du graphique",
"labels.title": "Titre du graphique",
"labels.xAxis": "Libellé de l'axe X",
"labels.yAxis": "Libellé de l'axe Y",
"labels.titleSettings": "Paramètres du titre",
"labels.fontFamily": "Police du titre",
"labels.fontSize": "Taille de police du titre",
"labels.showTitle": "Afficher le titre du graphique",

// PANNEAU AXES
"axes.heading": "Axes",
"axes.xSection": "Axe X",
"axes.ySection": "Axe Y",
"axes.xFontFamily": "Police de l'axe X",
"axes.xFontSize": "Taille de police de l'axe X",
"axes.yMin": "Minimum de l'axe Y",
"axes.yMax": "Maximum de l'axe Y",
"axes.yRangeAria": "Aide sur l'échelle de l'axe Y",
"axes.yRangeHelp": "Laissez les valeurs minimale et maximale vides pour que le graphique détermine automatiquement son échelle.",
"axes.yFontFamily": "Police de l'axe Y",
"axes.yFontSize": "Taille de police de l'axe Y",
"axes.yFormat": "Format de l'axe Y",
"axes.formatAuto": "Auto",
"axes.formatDecimal": "Décimal",
"axes.formatInteger": "Entier",
"axes.formatPercent": "Pourcentage",
"axes.decimals": "Nombre de décimales de l'axe Y",
"axes.decimalsAuto": "Auto",
"axes.dp0": "0 décimale",
"axes.dp1": "1 décimale",
"axes.dp2": "2 décimales",
"axes.dp3": "3 décimales",
"axes.tickInterval": "Intervalle des graduations de l'axe Y",
"axes.tickAria": "Aide sur l'intervalle des graduations",
"axes.tickHelp": "Laissez ce champ vide pour un espacement automatique. Pour les pourcentages, indiquez des points de pourcentage : par exemple, 1 correspond à 1 %, et 0.5 correspond à 0,5 %.",

// PANNEAU PÉRIODE DE RÉFÉRENCE
"baseline.heading": "Données de référence",
"baseline.helpAria": "Aide sur la période de référence",
"baseline.help": "Une période de référence est une période initiale utilisée pour comprendre le fonctionnement du processus avant les observations suivantes. Indiquez le nombre de points à utiliser pour établir cette référence. SimpleSPC utilise ces points pour calculer la ligne centrale initiale et, le cas échéant, les limites de contrôle. Les observations suivantes restent visibles pour permettre la comparaison. Laissez ce champ vide si vous n'avez pas besoin d'une période de référence initiale distincte. Choisissez cette période en fonction de votre connaissance du processus, et non simplement pour éliminer des signaux gênants.",
"baseline.useFirst": "Utiliser les",
"baseline.pointsAsBaseline": "premiers points comme période de référence",
"baseline.blankHint": "(laisser vide si aucune période de référence)",
"baseline.targetHeading": "Objectif",
"baseline.targetHelpAria": "Aide pour définir un objectif",
"baseline.targetHelp": "Indiquez le niveau de performance que vous souhaitez atteindre, puis précisez si les valeurs supérieures ou égales, ou inférieures ou égales, à cet objectif sont souhaitables. Un objectif est un niveau de performance visé, et non une limite de contrôle statistique.",
"baseline.targetHint": "Indiquez votre valeur cible, puis choisissez si des valeurs plus élevées ou plus faibles sont préférables.",
"baseline.targetValue": "Valeur cible",
"baseline.atOrAbove": "Une valeur supérieure ou égale à l'objectif est favorable",
"baseline.atOrBelow": "Une valeur inférieure ou égale à l'objectif est favorable",
"baseline.hideTarget": "Masquer la ligne d'objectif",

"baseline.showTarget": "Afficher la ligne d'objectif",

// PANNEAU DES RÈGLES
"rules.heading": "Règles de détection",
"rules.explainer": "Explication des règles",

"rules.shift": "Longueur d'une série (décalage)",
"rules.shiftAria": "Aide sur la règle de décalage",
"rules.shiftHelp": "Nombre de points consécutifs du même côté de la ligne centrale nécessaires pour détecter un décalage. Valeur par défaut : 8.",

"rules.trend": "Longueur de la tendance",
"rules.trendAria": "Aide sur la règle de tendance",
"rules.trendHelp": "Nombre d'augmentations ou de diminutions consécutives nécessaires pour détecter une tendance. Valeur par défaut : 6. Ce paramètre est utilisé uniquement lorsque la règle de tendance est disponible et activée.",

"rules.advanced": "Options de détection avancées",
"rules.enableTrend": "Activer la règle de tendance",
"rules.enableTrendHelp": "Recherche une succession prolongée d'augmentations ou de diminutions consécutives. La longueur requise est définie par le paramètre Longueur de la tendance ci-dessus.",
"rules.advancedCaution": "Option avancée — peut augmenter le nombre de fausses alertes.",

"rules.rareLabel": "Règles de séries et de tendances pour les cartes T/G",
"rules.rareAria": "Aide sur les règles pour événements rares",
"rules.rareHelp": "Active la détection des séries et des tendances sur les cartes T et G pour événements rares. Ces types de cartes présentent naturellement des irrégularités ; les règles de détection de motifs doivent donc être utilisées avec prudence.",
"rules.rareCaution": "Option avancée — les cartes d'événements rares peuvent produire de faux signaux de séries ou de tendances.",

"rules.clamp": "Limiter la LIC à zéro",
"rules.clampAria": "Aide sur la limite inférieure de contrôle",
"rules.clampHelp": "Si la limite inférieure de contrôle calculée est négative, afficher la LIC à zéro à la place. Utile pour les mesures telles que les dénombrements ou les durées, qui ne peuvent pas être négatives.",

"rules.zoneIntro": "Règles de zones facultatives (désactivées par défaut). À utiliser uniquement pour les types de cartes qui les prennent en charge.",
"rules.zone23": "Appliquer la règle « 2 points sur 3 dans le tiers extérieur »",
"rules.zone45": "Appliquer la règle « 4 points sur 5 au-delà de 1 sigma »",
"rules.zoneCaution": "Les règles supplémentaires de détection de motifs peuvent augmenter les fausses alertes sur les cartes XmR / individuelles. Elles sont désactivées par défaut et ne doivent être utilisées qu'en comprenant ce compromis.",

"rules.mrHeading": "Calcul des limites XmR",
"rules.mrAria": "Aide sur le filtrage des étendues mobiles XmR",
"rules.mrHelp": "Une variation très importante entre deux points consécutifs peut rendre les limites de contrôle XmR trop larges. Lorsque cette option est activée, SimpleSPC conserve tous les points et toutes les étendues mobiles sur le graphique, mais n'utilise pas les étendues mobiles anormalement grandes pour estimer la variation habituelle. La méthode de Nelson est utilisée : les étendues mobiles supérieures à 3,5 fois la moyenne initiale des étendues mobiles sont exclues de l'estimation de la variation.",
"rules.mrScreen": "Exclure les variations exceptionnellement grandes du calcul des limites",

"rules.conservative": "Cette carte utilise un ensemble de règles prudent, conçu pour permettre une interprétation fiable.",
"rules.flag": "Signaler les causes spéciales sur le graphique",
"rules.colourBlind": "Couleurs et symboles adaptés aux personnes daltoniennes",

// INFOBULLES DYNAMIQUES
"rules.disabledTrend": "La règle de tendance n'est pas proposée pour ce type de carte.",
"rules.enableTrendFirst": "Cochez « Activer la règle de tendance » pour utiliser ce paramètre.",
"rules.enableRareFirst": "Activez les règles de séries et de tendances pour ce type de carte afin d'utiliser ce paramètre.",
"rules.rareOnly": "Utilisé uniquement pour les cartes T et G.",
"rules.zoneBlocked": "Les règles de zones ne sont pas disponibles pour ce type de carte, car elles pourraient produire des alertes trompeuses.",

// CONFIRMATIONS POUR LES ÉVÉNEMENTS RARES
"rules.rareConfirmImmediate": "Les cartes T et G présentent naturellement des irrégularités.\n\nLes règles de séries et de tendances peuvent produire de faux signaux sur ce type de carte et sont désactivées par défaut.\n\nN'activez ces règles que si vous comprenez ce compromis.",

"rules.rareConfirmOnce": "Option avancée : règles de séries et de tendances sur les cartes {chart}\n\nCes motifs apparaissent souvent par hasard sur les cartes T/G et peuvent augmenter les fausses alertes.\n\nActiver malgré tout ?",

// FENÊTRE D'EXPLICATION DES RÈGLES
"explainer.title": "Explication des règles SPC",
"explainer.close": "Fermer",
"explainer.closeAria": "Fermer l'explication des règles",
"explainer.subtitle": "{chart} : cet outil utilise des règles prudentes, adaptées au type de carte, afin de limiter les fausses alertes.",

// NOMS DES CARTES
"explainer.chart.run": "Graphique chronologique",
"explainer.chart.xmr": "Carte XmR",
"explainer.chart.c": "Carte C",
"explainer.chart.p": "Carte P",
"explainer.chart.pprime": "Carte P′",
"explainer.chart.u": "Carte U",
"explainer.chart.uprime": "Carte U′",
"explainer.chart.xbars": "Carte X̄–S",
"explainer.chart.t": "Carte T",
"explainer.chart.g": "Carte G",
"explainer.chart.pareto": "Diagramme de Pareto",

// POURQUOI LES RÈGLES DIFFÈRENT
"explainer.why.attribute": "Les cartes pour données attributaires utilisent ici un ensemble de règles prudent. Les règles de zones sont désactivées, car elles peuvent produire des alertes trompeuses, notamment lorsque les limites varient ou que les dénombrements sont faibles.",
"explainer.why.rare": "Les cartes d'événements rares présentent naturellement des irrégularités. Les règles de séries et de tendances peuvent produire de faux signaux ; leur activation est donc réservée aux options avancées et accompagnée d'un avertissement.",
"explainer.why.xmr": "Les cartes XmR suivent des mesures individuelles dans le temps. SimpleSPC conserve volontairement une interprétation simple par défaut, tout en proposant des règles supplémentaires de détection de motifs et des options de calcul pour les utilisateurs qui en ont besoin.",
"explainer.why.xbars": "Les cartes X̄–S constituent la famille de cartes pour laquelle des règles avancées sont les plus justifiables, mais les règles activées par défaut restent prudentes pour limiter les fausses alertes.",
"explainer.why.run": "Les graphiques chronologiques n'utilisent pas de limites de contrôle. L'outil privilégie par défaut une interprétation simple et prudente.",
"explainer.why.pareto": "Un diagramme de Pareto n'est pas une carte de contrôle ordonnée dans le temps : les règles de détection des signaux SPC ne s'y appliquent donc pas. Il classe les catégories pour orienter les investigations et les actions d'amélioration.",
"explainer.why.general": "Cet outil privilégie par défaut un ensemble de règles simple et prudent afin de limiter les fausses alertes.",

// RÈGLES APPLICABLES
"explainer.checksHeading": "Ce que cette carte vérifie par défaut",
"explainer.rule.beyond": "Points au-delà des limites",
"explainer.rule.run": "Règle des séries",
"explainer.rule.trend": "Règle de tendance",
"explainer.rule.zone23": "2 points sur 3 dans le tiers extérieur",
"explainer.rule.zone45": "4 points sur 5 au-delà de 1 sigma",
"explainer.status.on": "appliquée par défaut.",
"explainer.status.optional": "disponible dans les options avancées, désactivée par défaut.",
"explainer.status.warn": "réservée aux options avancées, désactivée par défaut, avec un avertissement avant utilisation.",
"explainer.status.unavailable": "non proposée pour ce type de carte.",

// POURQUOI CERTAINES RÈGLES SONT LIMITÉES
"explainer.limitedHeading": "Pourquoi certaines règles sont limitées",
"explainer.limited1": "Cet outil est conçu pour favoriser une interprétation fiable dans les secteurs de la santé et des services publics. Il privilégie un nombre réduit de signaux plus fiables plutôt qu'un ensemble de règles produisant davantage de fausses alertes.",
"explainer.limited2": "Un « signal » peut entraîner des réunions, des inquiétudes, des investigations ou des actions. Pour réduire les efforts inutiles et la fatigue liée aux alertes, certaines règles sont réservées aux types de cartes pour lesquels elles sont les plus fiables.",

// DÉFINITIONS DES RÈGLES
"explainer.definitionsHeading": "Définitions simples des règles",
"explainer.definition.beyond": "<strong>Points au-delà des limites :</strong> un point au-dessus de la limite supérieure ou au-dessous de la limite inférieure de contrôle.",
"explainer.definition.run": "<strong>Règle des séries :</strong> une suite prolongée de points du même côté de la ligne centrale.",
"explainer.definition.trend": "<strong>Règle de tendance :</strong> une succession prolongée d'augmentations ou de diminutions consécutives.",
"explainer.definition.zone": "<strong>Règles de zones :</strong> règles supplémentaires fondées sur la distance des points par rapport à la ligne centrale. Leur applicabilité est plus limitée ; elles ne sont proposées que pour les types de cartes où leur usage est le mieux justifié.",

// EXPLICATION TECHNIQUE DES CARTES XMR
"explainer.xmr.heading": "Comment sont calculées les limites XmR",
"explainer.xmr.intro": "Les cartes XmR estiment la variation habituelle à partir des différences entre les points consécutifs, appelées <strong>étendues mobiles</strong>.",
"explainer.xmr.largeJump": "Un écart très important entre deux points peut surestimer la variation et produire des limites de contrôle trop larges. Par défaut, FreeSPC écarte les étendues mobiles exceptionnellement grandes de l'estimation utilisée pour calculer les limites XmR. Les points et les étendues mobiles d'origine restent visibles sur la carte.",
"explainer.xmr.technical": "<strong>Détail technique :</strong> FreeSPC utilise la méthode de Nelson. Les étendues mobiles supérieures à 3,5 fois la moyenne initiale des étendues mobiles ne sont pas utilisées pour estimer la variation habituelle.",
"explainer.xmr.current": "<strong>Paramètre actuel :</strong> le filtrage des étendues mobiles est <strong>{state}</strong>. Vous pouvez modifier ce réglage dans les <em>Options de détection avancées</em>.",
"explainer.xmr.split": "<strong>Que se passe-t-il lorsque vous ajoutez un changement de phase ?</strong> FreeSPC traite les données avant et après ce point comme deux périodes distinctes du processus. Chaque période possède sa propre moyenne et ses propres limites de contrôle ; l'étendue mobile qui franchit la séparation n'est pas utilisée.",
"explainer.xmr.minimum": "FreeSPC nécessite au moins <strong>{points} points</strong> pour interpréter les limites XmR d'une période. Si un changement de phase crée une période plus courte, les limites sont présentées comme provisoires et FreeSPC ne conclut pas encore que cette période présente uniquement une variation habituelle ou des causes spéciales.",
"explainer.on": "ACTIVÉ",
"explainer.off": "DÉSACTIVÉ"






});
