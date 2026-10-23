
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
"explainer.off": "DÉSACTIVÉ",

// ZONE DES RÉSULTATS
"results.chartHeading": "Graphique :",
"results.chartHelpAria": "Aide sur le menu contextuel du graphique",
"results.chartHelp": "Faites un clic droit sur le graphique pour accéder à des options supplémentaires. Selon le type de carte, vous pouvez ajouter des annotations, créer ou supprimer des changements de phase, exclure ou réintégrer des points étudiés dans le calcul des limites, gérer les libellés des lignes centrales et copier ou exporter votre graphique. Certaines actions nécessitent un clic droit directement sur un point de données.",

// GUIDE DE DÉMARRAGE
"results.quickStart": "Démarrage rapide :",
"results.quick1": "Importez un fichier CSV ou Excel, ou ouvrez l'éditeur de données pour coller ou saisir des données",
"results.quick2Before": "Si nécessaire, corrigez les données ou supprimez les lignes et colonnes inutiles dans l'éditeur, puis cliquez sur",
"results.quickApply": "Appliquer les données",
"results.quick3": "Choisissez le type de carte et vérifiez les colonnes X et Y",
"results.quick4Before": "Cliquez sur",
"results.quick4After": "si les modifications ne sont pas prises en compte automatiquement",
"results.quick5": "Ajoutez des titres, des libellés, des objectifs, des changements de phase ou des annotations selon vos besoins",
"results.quick6Before": "Faites un clic droit sur le graphique pour accéder aux options supplémentaires et utilisez",
"results.quick6After": "si vous souhaitez reprendre votre travail ultérieurement",

// CONTRÔLES DYNAMIQUES
"results.hideControls": "Masquer les paramètres",
"results.showControls": "Afficher les paramètres",
"results.needsRecalc": "Modifications enregistrées — cliquez sur Générer / Recalculer",

// TITRES DES CARTES SECONDAIRES
"results.mrHeading": "Carte des étendues mobiles :",
"results.sHeading": "Carte S :",

"results.mrCanvasTitle": "Étendues mobiles (MR)",

// LÉGENDES DES CARTES
"visual.value": "Valeur",
"visual.centreLine": "Ligne centrale",
"visual.median": "Médiane",
"visual.ucl": "LSC",
"visual.lcl": "LIC",
"visual.target": "Objectif",
"visual.xbar": "X̄",
"visual.count": "Effectif",
"visual.cumulative": "Pourcentage cumulé",
"visual.movingRange": "Étendue mobile",
"visual.mrAverage": "Moyenne des étendues mobiles",
"visual.mrUcl": "LSC des étendues mobiles",
"visual.mrAveragePeriod": "Moyenne des étendues mobiles (période {period})",
"visual.mrUclPeriod": "LSC des étendues mobiles (période {period})",
"visual.omittedExcluded": "Non utilisée : adjacente à un point exclu",
"visual.omittedScreened": "Non utilisée : étendue mobile exceptionnellement grande",

// INFOBULLES DE LA CARTE DES ÉTENDUES MOBILES
"visual.mrTooltip": "Étendue mobile : {value}",
"visual.mrExcludedTooltip": "Non utilisée pour estimer la variation : {value} (adjacente à un point exclu)",
"visual.mrScreenedTooltip": "Non utilisée pour estimer la variation : {value} (étendue mobile exceptionnellement grande, écartée du calcul)",

// LIBELLÉS SUPPLÉMENTAIRES XMR
"visual.mean": "Moyenne",
"visual.ucl3sigma": "LSC (3σ)",
"visual.lcl3sigma": "LIC (3σ)",

// INFOBULLES DU DIAGRAMME DE PARETO
"visual.pareto.count": "Effectif : {count} ({share} % du total)",
"visual.pareto.cumulative": "Pourcentage cumulé : {percent} %",

"visual.pareto.through": "Pourcentage cumulé jusqu'à cette catégorie : {percent} %",

// TITRES AUTOMATIQUES DES CARTES
"auto.title.run": "Graphique chronologique",
"auto.title.xmr": "Carte X (valeurs individuelles)",
"auto.title.c": "Carte C",
"auto.title.p": "Carte P",
"auto.title.pprime": "Carte P′",
"auto.title.u": "Carte U",
"auto.title.uprime": "Carte U′",
"auto.title.t": "Carte T",
"auto.title.tGaps": "Carte T (intervalles)",
"auto.title.g": "Carte G",
"auto.title.pareto": "Diagramme de Pareto",
"auto.title.xbar": "Carte X̄",
"auto.title.s": "Carte S",

// LIBELLÉS AUTOMATIQUES DES AXES
"auto.axis.date": "Date",
"auto.axis.value": "Valeur",
"auto.axis.category": "Catégorie",
"auto.axis.count": "Effectif",
"auto.axis.subgroup": "Sous-groupe",
"auto.axis.proportion": "Proportion / %",
"auto.axis.rate": "Taux par opportunité",
"auto.axis.rateN": "Taux pour {n} opportunités",
"auto.axis.opportunities": "Opportunités entre événements",
"auto.axis.time": "Temps entre événements",
"auto.axis.timeDays": "Temps entre événements (jours)",
"auto.axis.cumulative": "Pourcentage cumulé",
"auto.axis.s": "S",

// INTRODUCTIONS DES RÉSUMÉS STATISTIQUES
"summary.heading": "Résumé",
"summary.xbarsHeading": "Résumé de la carte X̄–S",
"summary.chart.run": "Graphique chronologique",
"summary.chart.xmr": "Carte XmR",

"summary.totalPoints": "Nombre total de points :",
"summary.totalSubgroups": "Nombre total de sous-groupes :",
"summary.divided": "Le graphique est divisé en",
"summary.period": "période",
"summary.periods": "périodes",
"summary.basedBaseline": "(selon la période de référence et les éventuels changements de phase).",
"summary.basedSplits": "(selon les changements de phase définis).",

"summary.rateIntro": "Taux affichés pour",
"summary.opportunity": "opportunité",
"summary.opportunities": "opportunités",

"summary.note": "Remarque :",
"summary.rareNote": "Les cartes d'événements rares présentent souvent des distributions asymétriques ; les limites de contrôle peuvent donc ne pas être symétriques comme celles d'une carte XmR.",

"summary.paretoOccurrences": "Nombre total d'occurrences :",
"summary.paretoAcross": "réparties dans",
"summary.paretoCategories": "catégories.",

// TITRES DES PÉRIODES ET COUVERTURE
"summary.periodLabel": "Période",
"summary.singlePeriod": "Période unique",
"summary.initialSegment": "(segment initial / période de référence)",
"summary.of": "sur",
"summary.coverage": "Période couverte :",
"summary.point": "point",
"summary.points": "points",
"summary.subgroups": "sous-groupes",
"summary.to": "à",

// MESURES STATISTIQUES
"summary.mean": "Moyenne",
"summary.medianPeriod": "Médiane (cette période)",
"summary.centreLine": "Ligne centrale",
"summary.controlLimits": "limites de contrôle",
"summary.lowerLimit": "LIC",
"summary.upperLimit": "LSC",
"summary.range": "intervalle",
"summary.estimatedSigma": "σ estimé (à partir des étendues mobiles)",
"summary.averageMr": "moyenne des étendues mobiles",
"summary.centreGap": "Ligne centrale (intervalle moyen entre événements)",
"summary.centreOpportunities": "Ligne centrale (nombre moyen d'opportunités)",
"summary.upperLimitLabel": "limite supérieure",
"summary.centreAndLimits": "Ligne centrale et limites",
"summary.notAvailable": "(non disponible)",

// MESURES DES CARTES X-BAR / S
"summary.subgroupSize": "Taille du sous-groupe",
"summary.measurement": "mesure",
"summary.measurements": "mesures",
"summary.perSubgroup": "par sous-groupe",
"summary.variesSubgroup": "(variable selon le sous-groupe)",
"summary.notStated": "non précisé",
"summary.xbarCentre": "Ligne centrale de la carte X̄",
"summary.sCentre": "Ligne centrale de la carte S",

"summary.estimatedProcessSd": "Écart-type estimé du processus",

// EXPLICATIONS DE LA PÉRIODE DE RÉFÉRENCE X-BAR / S
"summary.xbars.baseline": "Période de référence",
"summary.xbars.firstOne": "le premier sous-groupe (1)",
"summary.xbars.firstMany": "les",
"summary.xbars.subgroupsMany": "premiers sous-groupes",
"summary.xbars.windowOne": "définit la période utilisée pour le calcul des limites ;",
"summary.xbars.windowMany": "définissent la période utilisée pour le calcul des limites ;",
"summary.xbars.usedEstimate": "sous-groupes pris en compte pour estimer les limites",
"summary.xbars.excludedOne": "exclu",
"summary.xbars.excludedMany": "exclus",

// EXPLICATIONS DES LIMITES DE CONTRÔLE X-BAR / S
"summary.xbars.limitsVary": "les limites de contrôle varient selon la taille du sous-groupe",
"summary.xbars.sLimitsVary": "La valeur attendue de S et ses limites de contrôle varient selon la taille du sous-groupe.",
"summary.xbars.centreRange": "Intervalle des lignes centrales",
"summary.xbars.lclRange": "Intervalle des LIC",
"summary.xbars.uclRange": "Intervalle des LSC",
"summary.xbars.limits": "limites",

// ÉTAT DES CARTES
"summary.xbars.xChart": "Carte X̄",
"summary.xbars.sChart": "Carte S",
"summary.xbars.interpretation": "Interprétation",
"summary.xbars.latestPeriod": "Dernière période",
"summary.xbars.stableX": "stable (aucun signal clair de changement dans les moyennes des sous-groupes).",
"summary.xbars.stableS": "stable (aucun signal clair de changement dans la variation au sein des sous-groupes).",
"summary.xbars.signals": "signal(s) :",
"summary.xbars.signalDetected": "un ou plusieurs signaux de cause spéciale ont été détectés.",

// INTERPRÉTATIONS COMBINÉES PAR PÉRIODE
"summary.xbars.bothStable": "Les moyennes des sous-groupes (X̄) et la variation au sein des sous-groupes (S) semblent toutes deux stables pendant cette période.",
"summary.xbars.xUnstable": "Les moyennes des sous-groupes (X̄) présentent des signaux de cause spéciale, tandis que la variation au sein des sous-groupes (S) semble stable.",
"summary.xbars.sUnstable": "Les moyennes des sous-groupes (X̄) semblent stables, tandis que la variation au sein des sous-groupes (S) présente des signaux de cause spéciale.",
"summary.xbars.bothUnstable": "Les moyennes des sous-groupes (X̄) et la variation au sein des sous-groupes (S) présentent toutes deux des signaux de cause spéciale.",

// ÉVALUATION DE LA DERNIÈRE PÉRIODE
"summary.xbars.latestBothStable": "La dernière période semble stable sur les deux parties de la carte X̄–S. Cela signifie que le niveau actuel du processus et la variation au sein des sous-groupes présentent un comportement globalement cohérent selon les règles de détection actuellement sélectionnées.",
"summary.xbars.latestXUnstable": "La dernière période présente un changement possible dans les moyennes des sous-groupes, tandis que la variation au sein des sous-groupes reste stable.",
"summary.xbars.latestSUnstable": "La dernière période présente un changement possible dans la variation au sein des sous-groupes, tandis que les moyennes des sous-groupes restent stables.",
"summary.xbars.latestBothUnstable": "La dernière période présente des signes possibles de variation due à des causes spéciales, aussi bien dans les moyennes des sous-groupes que dans la variation au sein des sous-groupes.",

// GRAPHIQUE CHRONOLOGIQUE — RÉFÉRENCE
"summary.run.baselineFirst": "Période de référence : les",
"summary.run.baselinePoints": "premiers points sont utilisés pour calculer la médiane.",
"summary.run.baselineAll": "Période de référence : tous les points de cette période sont utilisés pour calculer la médiane.",

// GRAPHIQUE CHRONOLOGIQUE — SIGNAUX
"summary.run.shiftBefore": "un décalage prolongé (≥",
"summary.run.shiftAfter": "points du même côté de la médiane)",
"summary.run.trendBefore": "une tendance prolongée (≥",
"summary.run.trendAfter": "points en augmentation ou en diminution)",
"summary.run.outlier": "une valeur exceptionnellement éloignée (point « astronomique »)",

// GRAPHIQUE CHRONOLOGIQUE — INTERPRÉTATION
"summary.run.noSignals": "Aucun signal clair de cause spéciale n'a été détecté pendant cette période (pas de décalage prolongé, de tendance prolongée ni de valeur exceptionnellement éloignée). Ce comportement est compatible avec une variation habituelle, mais doit toujours être interprété dans son contexte.",
"summary.run.hasSignals": "Cette période présente des signaux de cause spéciale :",

// GRAPHIQUE CHRONOLOGIQUE — POINTS À EXAMINER
"summary.run.where": "Points à examiner :",
"summary.run.whereShift": "décalage autour des points",
"summary.run.whereTrend": "tendance autour des points",
"summary.run.whereOutlierOne": "valeur exceptionnellement éloignée au point",
"summary.run.whereOutlierMany": "valeurs exceptionnellement éloignées aux points",
"summary.run.increasing": "croissante",
"summary.run.decreasing": "décroissante",

// GRAPHIQUE CHRONOLOGIQUE — CHANGEMENTS DE PHASE
"summary.run.splitNote": "Chaque période est résumée séparément, car les changements de phase suggèrent que le processus a pu évoluer au fil du temps.",

// XMR — DESCRIPTIONS DES SIGNAUX
"summary.xmr.signalBeyond": "un ou plusieurs points se situent au-delà des limites de contrôle",
"summary.xmr.signalShift": "un décalage prolongé (plusieurs points du même côté de la moyenne)",
"summary.xmr.signalTrend": "une tendance prolongée (augmentation ou diminution continue)",
"summary.xmr.signalOutlier": "une valeur exceptionnellement éloignée (point « astronomique »)",

// XMR — PÉRIODE DE RÉFÉRENCE
"summary.xmr.baselineWindowStart": "Période de référence : fenêtre initiale de",
"summary.xmr.baselineWindowOne": "point servant de référence ;",
"summary.xmr.baselineWindowMany": "points servant de référence ;",
"summary.xmr.baselineUsed": "pris en compte pour calculer la moyenne et les limites",
"summary.xmr.excludedOne": "point exclu du calcul",
"summary.xmr.excludedMany": "points exclus du calcul",
"summary.xmr.baselineAll": "Période de référence : tous les points de cette période ont été utilisés pour calculer la moyenne et les limites.",

// XMR — EXCLUSIONS MANUELLES
"summary.xmr.exclusionHeading": "Ajustement lié aux points exclus",
"summary.xmr.adjacentOne": "étendue mobile adjacente n'a pas été utilisée pour estimer la variation, car elle comprend un point exclu.",
"summary.xmr.adjacentMany": "étendues mobiles adjacentes n'ont pas été utilisées pour estimer la variation, car elles comprennent un point exclu.",
"summary.xmr.noArtificialMR": "SimpleSPC ne crée pas d'étendue mobile artificielle de part et d'autre d'un point exclu.",

// XMR — FILTRAGE DE NELSON
"summary.xmr.limitCalculation": "Calcul des limites XmR",
"summary.xmr.screenedOne": "variation exceptionnellement grande entre deux points consécutifs n'a pas été utilisée pour estimer la variation habituelle. Tous les points restent visibles sur la carte.",
"summary.xmr.screenedMany": "variations exceptionnellement grandes entre deux points consécutifs n'ont pas été utilisées pour estimer la variation habituelle. Tous les points restent visibles sur la carte.",
"summary.xmr.nelsonUsed": "Le filtrage des étendues mobiles selon la méthode de Nelson a été appliqué :",
"summary.xmr.initialMR": "moyenne initiale des étendues mobiles",
"summary.xmr.threshold": "seuil de filtrage",
"summary.xmr.revisedMR": "moyenne révisée des étendues mobiles",
"summary.xmr.screeningOff": "Le filtrage des étendues mobiles est désactivé ; toutes les variations entre points consécutifs de la période de référence ont donc été utilisées pour estimer la variation habituelle.",

// XMR — LIMITES PROVISOIRES
"summary.xmr.provisionalHeading": "Limites provisoires",
"summary.xmr.provisionalOnly": "seulement",
"summary.xmr.provisionalOne": "point a été utilisé pour estimer les limites de cette période.",
"summary.xmr.provisionalMany": "points ont été utilisés pour estimer les limites de cette période.",
"summary.xmr.minimumIntro": "SimpleSPC nécessite au moins",
"summary.xmr.minimumEnd": "points avant de pouvoir interpréter les signaux de cause spéciale ou estimer la capabilité du processus.",
"summary.xmr.provisionalContext": "Les limites affichées sont fournies uniquement à titre indicatif.",

// XMR — OBJECTIFS
"summary.xmr.target": "Objectif",
"summary.xmr.aboveBetter": "une valeur supérieure ou égale est préférable",
"summary.xmr.belowBetter": "une valeur inférieure ou égale est préférable",
"summary.xmr.targetCoverage": "des points de cette période atteignent l'objectif",
"summary.xmr.coverageUnavailable": "La proportion de points atteignant l'objectif n'a pas été calculée pour cette période.",

// XMR — INTERPRÉTATION
"summary.xmr.notAssessed": "Évaluation non encore possible : davantage de données sont nécessaires avant de pouvoir déterminer si cette période présente une variation habituelle ou des signaux de cause spéciale.",
"summary.xmr.noSignals": "Aucun signal clair de cause spéciale n'a été détecté pendant cette période. Le comportement observé est compatible avec une variation habituelle, mais doit toujours être interprété dans son contexte clinique.",
"summary.xmr.hasSignals": "Cette période présente des signaux de cause spéciale :",
"summary.xmr.where": "Points à examiner",
"summary.xmr.outsideOne": "point au-delà des limites au rang",
"summary.xmr.outsideMany": "points au-delà des limites aux rangs",
"summary.xmr.outlierOne": "valeur exceptionnellement éloignée au rang",
"summary.xmr.outlierMany": "valeurs exceptionnellement éloignées aux rangs",

// XMR — CAPABILITÉ
"summary.xmr.estimatedCapability": "Capabilité estimée (cette période)",
"summary.xmr.ifStable": "si le processus reste stable, environ",
"summary.xmr.futureMeet": "des points futurs devraient atteindre l'objectif.",
"summary.xmr.capability": "Capabilité",
"summary.xmr.capabilityUnreliable": "un objectif a été défini, mais des signaux de cause spéciale sont présents pendant cette période ; toute estimation de capabilité serait donc peu fiable.",
"summary.xmr.comparePeriods": "La comparaison des moyennes, des limites et des performances par rapport à l'objectif entre les périodes peut aider à déterminer si le processus a changé après des interventions.",

// XMR — INDICATEUR DE CAPABILITÉ
"summary.xmr.badgeHeading": "CAPABILITÉ DU PROCESSUS (dernière période)",
"summary.xmr.badgeNote": "(Probabilité estimée d'atteindre l'objectif pendant la dernière période, sous l'hypothèse d'un processus stable et d'une distribution approximativement normale.)",
"summary.xmr.badgeWarning": "Signaux de cause spéciale détectés pendant la dernière période :",
"summary.xmr.badgeUnreliable": "les estimations de capabilité peuvent être peu fiables tant que ces signaux ne sont pas compris.",
"summary.xmr.badgeAction": "Cherchez à comprendre et à traiter ces causes avant de vous appuyer sur des estimations de capabilité.",

"summary.exclusions.heading": "Exclus du calcul des limites :",
"summary.exclusions.point": "point",
"summary.exclusions.points": "points",
"summary.exclusions.subgroup": "sous-groupe",
"summary.exclusions.subgroups": "sous-groupes",
"summary.exclusions.pointVisible": "Le point exclu reste visible sur le graphique.",
"summary.exclusions.pointsVisible": "Les points exclus restent visibles sur le graphique.",
"summary.exclusions.subgroupVisible": "Le sous-groupe exclu reste visible sur le graphique.",
"summary.exclusions.subgroupsVisible": "Les sous-groupes exclus restent visibles sur le graphique.",

// CARTES AUX ATTRIBUTS — PÉRIODE DE RÉFÉRENCE
"summary.attr.baselineTitle": "Période de référence pour cette période",
"summary.attr.firstOne": "le premier point (1) définit la fenêtre de référence ;",
"summary.attr.firstBefore": "les",
"summary.attr.firstAfter": "premiers points définissent la fenêtre de référence ;",
"summary.attr.used": "utilisés pour calculer la ligne centrale et les limites",
"summary.attr.excludedOne": "exclu du calcul",
"summary.attr.excludedMany": "exclus du calcul",

// CARTES AUX ATTRIBUTS — SIGNAUX ET INTERPRÉTATION
"summary.attr.signals": "Signaux",
"summary.attr.stable": "Selon les règles sélectionnées, aucun signal clair de cause spéciale n'a été détecté pendant cette période. Cela suggère un comportement compatible avec une variation habituelle (de cause commune).",
"summary.attr.unstable": "Selon les règles sélectionnées, des signaux de cause spéciale ont été détectés pendant cette période (un comportement peu susceptible de s'expliquer uniquement par la variation habituelle).",

// CARTES AUX ATTRIBUTS — PRÉCAUTIONS STATISTIQUES
"summary.attr.caution": "Ces règles sont des indications, et non des réponses définitives. Interprétez les résultats en tenant compte du contexte local (modifications du processus, effectifs, demande, définitions ou codage) et des hypothèses de base des cartes SPC, notamment la comparabilité des mesures et des opportunités d'observation au fil du temps.",

// CARTES AUX ATTRIBUTS — SIGNAUX INDIVIDUELS
"summary.attr.signalAbove": "Un ou plusieurs points dépassent la limite supérieure de contrôle",
"summary.attr.signalBelow": "Un ou plusieurs points se situent sous la limite inférieure de contrôle",

"summary.attr.shiftName": "Décalage",
"summary.attr.shiftAbove": "points consécutifs au-dessus de la ligne centrale (de",
"summary.attr.shiftBelow": "points consécutifs au-dessous de la ligne centrale (de",

"summary.attr.trendName": "Tendance",
"summary.attr.trendIncreasing": "points en augmentation régulière (de",
"summary.attr.trendDecreasing": "points en diminution régulière (de",

"summary.attr.zone23": "Zone : 2 points sur 3 dans le tiers extérieur (de",
"summary.attr.zone45": "Zone : 4 points sur 5 au-delà de 1 sigma (de"





});
