# Suivi financier — état du projet (référence pour Claude)

**Version documentée : `v 2026.09.26.09.58`** — comparer à `APP_VERSION` (tête du `<script>` de `index.html`, format `v AAAA.MM.JJ.HH.MM`). Si différent, le code fait foi ; zones les plus mouvantes : §2, §4, §6.2–6.7, §7.2, §10.

**Fichiers** (tous uploadés ensemble sur GitHub Pages à chaque livraison, en zip avec arborescence) : `index.html` (application complète, vanilla JS, style Samsung One UI, mobile-first), `manifest.json`, `service-worker.js`, `icon-192.png`, `icon-512.png`, `icon-maskable-192.png`, `icon-maskable-512.png`, `apple-touch-icon.png`, plus le compagnon `google-apps-script.gs` (relais Drive, §10 — ne jamais y écrire la vraie clé : dépôt public). CDN : SheetJS 0.18.5 (`defer`), Google Fonts. Pas de backend ; seule sortie de données : la synchro vers le Google Drive de l'utilisateur via son propre script (§10).

**Usage** : visualisation/filtrage mobile d'un suivi de dépenses Excel saisi à la main ; import incrémental dédupliqué. Locale française (JJ/MM/AAAA, virgule décimale). Contexte couple : "Moi" / "Ma femme", scope Perso/Couple.

**Livraison** : `APP_VERSION` et `SW_VERSION` (service-worker.js) mis au même horodatage en dernière étape ; `node --check` sur le script extrait ; tests Playwright (Chromium headless, `serviceWorkers:'block'` dans le contexte sinon SheetJS passe par le cache SW et échappe à l'interception réseau ; seeding via `window.__SF_STORE`/`__SF_DATA.generateDemoData`).

---

## 1. Modèle de données

### 1.1 Colonnes d'import (Excel/CSV, mapping par nom d'en-tête normalisé, ordre libre)
`Date · Montant · Type · Catégorie · Personne · Compte · Compte cible · Groupe · Remarque`
- **Date** (`U.parseDateValue`) : cellule date, numéro de série, texte `JJ/MM/AAAA` ou `JJ/MM/AA` (pivot 68 : `00–68` → 20xx).
- **Type** (`U.normType`) : `Income/Expense/Transfer` ou français (`revenu`, `dép…`/`exp…`, `virement`/`trans…`), insensible casse/accents. Export en français (`Revenu/Dépense/Transfert`).
- **Montant** : signe porté par la valeur (jamais déduit du Type), virgule acceptée.

### 1.2 Transaction interne
```
{ id, date, dateISO, montant (signé), type, categorie, personne, compte, compteCible, groupe, remarque,
  groupeBase,   // groupe sans suffixe "(...)" — helper d'affichage/repli uniquement (§2.1)
  transform }   // { mode: "same"|"month"|"yearsplit", year, month }
```
`id` = hash du contenu (`buildTxnId`) → change à chaque modification d'un champ.

### 1.3 Date transformée (`computeTransform()`, depuis la Remarque)
- suffixe `AA-MM` → affectée à ce mois (`month`) ;
- suffixe `AAAA` → répartie sur les 12 mois (`yearsplit`) ; dans un contexte de période (Budget, Évolution "Mois transformé"), seule la part des mois de la période compte ;
- sinon `same`.
Calculé une fois au chargement/import, réutilisé tel quel par Budget (`computeGroupBudget`) et Évolution (`buildEvolutionIndex`).

---

## 2. Groupes de groupe / groupes de personnes

Plusieurs groupes datés (`Cliff Haven (23-01-08)`, `Cliff Haven (24-02)`) se rattachent à un groupe de groupe (`Ghana`). Règles toujours manuelles.

### 2.1 Règle d'or : le libellé complet est la clé, jamais le nom de base
- `findGroupAliasRule(base, full)` : règle exacte sur `full` d'abord ; repli sur une règle "générique" (`groupe = base`, sans `full`) seulement pour les occurrences sans règle propre. Aucune fenêtre de dates sur les règles.
- `getGroupKind` / `getGroupVoyageMonth` / `getGroupDateRange` : indexés sur le libellé complet (`editingFull || editingBase` côté UI).
- Une règle générique n'existe que si l'utilisateur coche "Appliquer à toutes les périodes de « X »" (§2.4).
- Ne jamais dériver un nom de base pour une recherche/comparaison de réglage : source historique d'échecs d'import silencieux.

### 2.1 bis Tri chronologique (`AGG.groupeHierarchyChrono(transactions)`)
Utilisé par : picker Groupe de la fiche transaction, filtre Groupe (Transactions, Santé, Revenus), vue Groupe d'Évolution, picker Groupe d'un ajustement Budget. Groupes de groupe du plus récent au plus ancien (selon la date de départ de leur enfant le plus ancien) ; enfants du plus ancien au plus récent ; groupes non rattachés à part (`flat`, alpha). Résultat marqué `preSorted:true` (les vues ne re-trient pas).
Le picker groupe principal/additionnel du Budget (`groupPickerHierarchy`) utilise `groupeHierarchyFull`, alphabétique.

**Date de départ** (`AGG.getGroupStartDate`) : 1) date de début configurée (§2.5) ; 2) suffixe du libellé (`AGG.parseGroupSuffixDate` : `AA-MM-JJ`, `AA-MM_MM` → 1ᵉʳ du 1ᵉʳ mois, `AA-MM`, `AAAA`) ; 3) transaction la plus ancienne du groupe (`AGG.groupStartDatesMap`).

### 2.2 Règles
```
groupAliasRules:  { groupe, superGroupe, full? }   // full = occurrence précise
personAliasRules: { personne, superPersonne }       // jamais de full
```
`getGroupeDisplay(t)` = super-groupe si règle, sinon `t.groupe` brut.

### 2.3 Application
- Filtres Groupe/Personne : liste imbriquée, `Filters.matches()` teste valeur résolue **et** brute.
- Liste Transactions : affiche toujours le groupe littéral (`t.groupe`).
- Regroupement "Compte" = compte réel ; "Type de compte" est un mode distinct.
- Évolution vue Groupe : sections = groupe de groupe, items = libellé complet.

### 2.4 Écrans "Gérer les groupes / personnes" — `createAliasManager(cfg)`
Instances `window.__SF_GROUPMGR` (`hasDates:true, showKind:true`) et `window.__SF_PERSONMGR` (`hasDates:false`).
- Une ligne par occurrence datée (libellé complet), icône de type ; scroll et `<details>` ouverts préservés entre rendus.
- Tap = édition **de cette occurrence** : Enregistrer crée une règle `full` ; case "Appliquer à toutes les périodes de « X »" = règle générique (absente si pas de suffixe → règle générique directe).
- Type Vie quotidienne/Voyage + **Mois voyage** (`Store.settings.groupVoyageMonth`, `AAAA-MM`) + dates de début/fin (§2.5), clé = libellé complet.
- Sélection multiple ("Déplacer vers…", règle générique en masse) : Personnes uniquement.
- `obsoleteRules()` : bouton "N règle(s) obsolète(s) · Nettoyer" si une règle ne vise plus aucune transaction (comparaison `U.normalize` sur la base, exacte sur `full`) ; confirmation listant les règles.
- **Piège** : les deux instances coexistent avec des ids identiques (`am-*`) → toute requête DOM doit être scopée au `sheet` local, jamais `document.getElementById`.

#### Mois calculé — `AGG.getCalcMonthYM(t)` (source unique)
1) groupe de type voyage **et** mois voyage renseigné → ce mois ; 2) `transform.mode==="month"` → mois transformé ; 3) mois réel. `yearsplit` n'est jamais éclaté ici. Utilisé par le regroupement et le filtre "Mois calculé".

### 2.5 Dates de début/fin → jours & moyennes (`Store.settings.groupDateRange = {[libellé]: {dateFrom, dateTo}}`)
Jours = `dateTo − dateFrom + 1` ; Revenus/Dépenses/Total par jour sur `t.groupe === clé`, calculés en direct dans l'écran d'édition ; message si date manquante ou jours ≤ 0.
Case "toujours aujourd'hui" : `dateTo = AGG.GROUP_DATE_TODAY_SENTINEL` (`"today"`), résolue par `AGG.resolveGroupDateToValue` au moment du calcul ; `dateTo` uniquement. Excel : écrit `Aujourd'hui`, relu (`aujourd'hui`/`today`, insensible casse/accents).

---

## 3. Types de compte et acronymes
- `Store.settings.accountTypes` `{compte: "Courant"|"Épargne"|libre}` (défaut Courant) : regroupement onglet Comptes et mode "Type de compte".
- `Store.settings.accountAcronyms` `{compte: "ACRO"}` (≤ 4 car.) ; `AGG.getAccountAcronym` sinon initiales (`deriveAccountAcronym`, ≤ 3). Affiché en bandeau vertical de chaque transaction.

---

## 4. Filtres (Transactions)
Menu en deux étapes (`MENU_ITEMS`) : **Période, Mois indiqué, Mois calculé, Type, Compte, Groupe, Catégorie, Personne** + recherche texte. AND entre champs, OR dans un champ.
- Filtres actifs en puces dans la sheet (`activeFiltersSectionHtml`, `F.removeChip`) ; badge = `Filters.activeCount()`.
- "Tout / Aucun" sur chaque section multi-choix sauf Période (`quickSelectRowHtml`/`wireQuickSelect`).
- Mois indiqué / calculé (`"AAAA-MM"`) groupés par année (`moisGroupedHtml`) : l'en-tête d'année sélectionne l'année entière (pas de champ année).
- Compte groupé par type (`comptesGroupedHtml`).
- Groupe/Personne : `nestedFilterHtml` → `U.hierarchyCheckListHtml`/`U.wireHierarchyCheckList` (composant partagé avec Santé et Revenus) ; Personne : en-tête "Sans type" pour les non rattachées (`NESTED_DIM_CONFIG.personnes.noGroupLabel`).
- Catégorie : Revenus A→Z puis Dépenses A→Z (`AGG.distinctValuesCategorieSorted`).

---

## 5. Décimales
`Store.settings.decimals` (0 par défaut, ou 2), pilule "0,00" dans Comptes, s'applique à Comptes et Transactions. Budget, Évolution, Récurrent, Revenus : toujours 0 décimale (`U.fmtWholeSigned`, `U.signClass`). Distinct de la pilule "0€" (masquer comptes à solde nul).

---

## 6. Onglets
Barre : **Comptes** (défaut), **Transactions**, **Budget**, **Évolution**, **Récurrent** ; bouton "⋯" (`#nav-more`, `openMoreMenu`, `MORE_TABS`) : **Revenus**, **Santé**.

### 6.1 Comptes
Solde total puis comptes par type (ordre `Store.settings.accountTypeOrder`, bouton Réorganiser `openAccountTypeReorderSheet`). Tap sur un compte → Transactions filtrées.
- **Copier pour Excel** (`buildAccountsClipboardText`) : TSV Compte/Solde + Total, repli `execCommand("copy")`, toujours tous les comptes.
- Pilule décimales (§5) ; pilule "0€" (`Store.settings.hideZeroAccounts`, `U.isZeroBalance`, local à l'appareil).
- **Soldes personnalisés** (`Store.settings.customBalances = [{title, accounts}]`, ≤ 2, titre vide = masqué).

### 6.2 Transactions
**Bouton flottant** (`#floatbar-wrap`, Transactions et Santé) : Affichage (`openDisplayMenu`), Loupe (bascule barre ↔ champ de recherche), Filtre, Plus (`openTxnNew`).

**11 regroupements** (`GROUPBY_LABEL`) : Compte, Type de compte, Jour, Mois indiqué, Mois calculé, Groupe, Période (= groupe de groupe), Catégorie, Personne, Groupe de personnes, Type. Groupe/Personne affichent le libellé brut ; la résolution vit dans Période/Groupe de personnes.

**Cartes de groupe** (`computeTxnGroups`/`renderTxnGroupCards`, un rendu pour tous les modes) : `<details class="txn-group-card">`, en-tête libellé + nombre + total. Repliées par défaut ; état par groupe dans `APP.state.txnOpenGroups` (réinitialisé par `resetTxnPagination`) ; "Tout déplier/replier" = `Store.settings.txnGroupsExpanded` (valeur par défaut, efface les exceptions). Scroll infini (`setupTxnInfiniteScroll`, sentinelle, +20 groupes) ; "Afficher plus" manuel au-delà de 10 transactions dans un groupe.

**Carte transaction** (`makeTxnItemEl`) : bandeau acronyme coloré par type ; ligne 1 titre (`txnTitle`, multi-lignes) ; ligne 2 Catégorie · Personne ou virement (`txnLine2`) ; ligne 3 date · groupe (`txnLine3`) ; montant aligné en haut. Lignes 2–3 tronquées.

**Fiche transaction** : `openTxnDetail(id, navIds?)` / `openTxnNew()` → `renderTxnDetail(id, seed, navList)`. Chaque champ s'édite au tap : Remarque en titre (`openInlineEditor` multiligne), bloc Détails (Montant, Date, Type, Compte ou Source/Cible pour un virement, Catégorie, Personne, Groupe). Pickers : Catégorie (`categoriePickerHierarchy`), Personne (`AGG.personneHierarchy`), Groupe (`groupeHierarchyChrono`), Type (3 options, `noEmpty`).
- Existante : chaque champ est enregistré immédiatement (`commit(patch)` → `DATA.rowToTxn`, id recalculé, suivi via `currentId`/`APP.state.detailTxnId`).
- Nouvelle / Dupliquer : brouillon, Enregistrer exige date + montant ; anti-collision d'id.
- Bas : ◀ Dupliquer Supprimer ▶ (existante) ; Annuler/Enregistrer (nouvelle).

**Navigation précédent/suivant** (seulement depuis la liste principale) : `navIds` = `flattenTxnGroupIds` de tout le résultat filtré ; instantané figé `nav` pendant que la fiche est ouverte, id mis à jour après chaque commit. Déclencheurs : boutons, swipe (seuil 60 px, horizontal ≥ 1,5× vertical), flèches clavier (si `APP.state.txnNavActive`, aucun champ focalisé, aucune autre fenêtre). `closeOverlay("overlay-detail")` remet `txnNavActive` à null et efface les handlers tactiles — `modal-detail` est partagé par de nombreuses fiches ; handlers assignés en propriétés (`.ontouchstart`), jamais `addEventListener`.

**`APP.confirmAction({title, text, confirmLabel, onConfirm})`** : confirmation générique (transaction, soin, ajustement, règle, ligne Revenus) ; `overlay-confirm` doit rester le dernier élément du shell.

### 6.3 Budget
**Modes** : "Amort. + ajouts" (`cfg.mode==="amorti"`) / "Ajouts" (`"brut_ajout"`). **Vue détail** (`cfg.detailView`) : 3 colonnes cliquables (`.bg-detail-cell`, `filterKind` `reel|amorti|ajout|reel_ajout`, `bgKindIncluded`) — Réel+ajout · Amorti · Total, ou Brut · Ajout · Total.
**Scope** Perso ×1 / Couple ×2 (transactions réelles seulement). **Vue** Total / Moyenne (÷ mois de la période affichée compris dans la borne d'amortissement `startMonth`→`endMonth` si définie — `amortBoundedMonths`).
**Période affichée** (`periodFrom`/`periodTo`, `null` = non définie → `budgetDataMonthBounds()`). Élargir la période ne contourne jamais la fenêtre d'amortissement ni les bornages (`extraGroupEffectiveMonths`, salaires).
Quatre boutons : Amortissement (début, lissage, catégories amorties, "Ne pas distinguer les voyages"), Ajustements, Comparaison, Comptes exclus.

- **En-tête figé** (comparaison ou vue détail) : `compareHeaderRowHtml`/`detailHeaderRowHtml`, `position:sticky;top:0`, enfant direct de la section qui contient toutes les cartes.
- **Carte de groupe** (principal / additionnels, même structure) : chevron = changer de groupe, ✈️ = fusion voyages de cette carte, ✕ pour un additionnel ; Total + Revenus / Dépenses / Voyages repliables (état commun `allSectionsOpen`).
- **Voyages** : un groupe de type Voyage alimente le bloc Voyages. Fusion dans les catégories si `budget.mergeVoyages` (global) OU `budget.mergeVoyagesGroups` (par carte) — `voyagesMergedForGroup`. La fusion ne change que la ligne de destination : mois résolu par **`voyageTxnMonths()`**, source unique partagée par `computeGroupBudget`, `openCategoryTxnSheet`, `collectGroupTypeItemsDetailed` (mois voyage forcé > transform > date réelle) ; une transaction voyage n'est jamais amortie.
- **Détail d'une ligne** : `openBudgetKindSheet(o)` + `BG_KIND_STYLE` (rendu commun des 4 feuilles : catégorie, voyage, section/total, mois) ; badges Réel/Amorti/Ajout filtrants ; bouton "Amortir" par transaction réelle ; yearsplit au prorata ("XX/12 mois").
- **Salaire** (catégorie normalisée "salaire") : tap = éditeur mensuel Moi / Ma femme / Total (`budget.salaryEntries[mois] = {moi, femme}`). `salaryMonthBreakdown` : Ma femme = Brut ; Moi = transactions réelles (Brut) + écart avec la valeur déclarée (Ajout) ; hors borne d'amortissement → ignoré. Alimente `realIncome`/`adjIncome`.
- **Comparaison** (`budget.compare = {enabled, mode:"period"|"predefined", periodFrom, periodTo, predefined}`) : colonnes Résultat · Comparé · Variation ; voyages en une ligne "Total voyages". `compareOnlyCategories()` affiche les catégories présentes seulement côté comparé (type via `classifyCategoryTypes`). Budget prédéfini = montant réel, complété en vue "Amort. + ajouts" par l'amortissement calculé de la même ligne (`predefinedRealValue`, `computeCompareContext.getAmortResult`, jamais pour un additionnel). Cellules "Comparé" cliquables (`.bg-cmp-cell`, `opts.periodOverride`) ; en prédéfini → éditeur.
- **Ajustements** (`openAdjustmentForm`) : mois, groupe (picker `adjustmentGroupHierarchy` = chrono + icône ✈️/🏠 via `kindFn`), catégorie (vide = ligne "Ajustements" des voyages), montant, remarque, scope Perso/Couple/Les deux, `montantViewOnly`, `amorti`. Correspondance carte : `adjustmentGroupMatches` (exact, nom de base, ou groupe précis résolu vers le groupe de groupe de la carte via `AGG.labelResolvesToSuperGroupe`) — jamais dans l'autre sens.
- **Fiches de réglages re-rendues en place** : sauvegarder/restaurer `modal.scrollTop` (et tout conteneur défilant interne, ex. `#pbf-list`) autour de chaque `innerHTML`.

### 6.4 Évolution
Vue (`evolution.view`) : Groupe / Catégorie / Personne / Compte ; base (`evolution.monthBasis`) : mois réel / transformé ; mode Normal/Cumulé (`evolution.cumulative`, local).
- `AGG.buildEvolutionIndex(dimension, monthBasis)` → `{sections:[{key,label,items:[{key,label,values: Map("YYYY-MM"→montant)}]}]}`, tri par montant (vue Groupe : chronologique).
- Compte : seule dimension incluant les virements (−source / +cible) ; le Total général les neutralise.
- Cumulé : `itemCumulativeYearValue`/`…MonthValue` (comparaison lexicographique des clés) ; la colonne Total reste la somme brute.
- Tableau `<table class="evo-table">` : Libellé + Total figés (`position:sticky`, largeurs px identiques, `box-sizing:border-box` obligatoire, `left` de la 2ᵉ colonne = largeur exacte de la 1ʳᵉ) — jamais de grille CSS (le sticky y échoue silencieusement). Colonnes par année, clic → 12 mois. `U.enableDragScroll` (seuil clic/glissé).
- Clic montant → `openEvolutionTxnSheet` (`txnMatchesDimTarget` + `txnMatchesPeriod`) ; ligne Total générale cliquable (`data-global`).
- Fusions de sections (`evolution.merges[dimension] = [{name, members}]`) : créées uniquement via JSON/synchro, supprimables en chip ; `__merged:true`.

### 6.5 Récurrent
Transactions à `transform.mode !== "same"`, par catégorie. `AGG.buildRecurringIndex(periodFrom, periodTo, groupFilter?)` : items = `U.baseRemarque(remarque)` (alpha), sections par montant absolu. `AA-MM` → colonne du mois ; `AAAA` → colonne spéciale après décembre (`item.yearValues`, `AGG.recurringColumns`). Seul Libellé est figé. Période (`recurrent.periodFrom/To`, local, défaut = tout via `ensureRecurrentPeriod`) sur le mois calculé. `recurrent.groupsExpanded`. Clic → `openRecurringTxnSheet` (`txnMatchesRecurringTarget` : mois / année spéciale / total).

### 6.6 Santé
Données, filtres et listes propres ; même présentation que Transactions.
```
santeItems[]: { id, dateSoin, beneficiaire, devise, montantEur, montantDevise,
  medecin, typePrestation, statut, statutDate,
  rembTiersMontant, rembTiersDate, rembSecuMontant, rembSecuDate, rembMutuelleMontant, rembMutuelleDate,
  groupe, mutuelle, partMoi, termine }
```
- `montantDevise` seulement si devise ≠ € (informatif, pas de conversion). `id` = `buildSanteId` (hors `termine`).
- `computeMaPart` = `montantEur × partMoi / 100` arrondi **au centime inférieur** (`Math.floor`).
- `termine` : pilule en tête de fiche, carte grisée, filtre + regroupement dédiés.
- Listes personnalisables `Store.settings.sante.lists` (`beneficiaires, devises, medecins, prestations, statuts, mutuelles`), ajout à la volée (`PICKER.openPicker({allowNew:true})`). Groupe = hiérarchie des Transactions. Choisir un Statut horodate `statutDate`.
- Filtres `__SF_SANTE_FILTERS` (sheet page unique, persistés `Store.settings.santeFilters`) ; comptages du filtre Groupe = ceux des transactions.
- Affichage (`SANTE_GROUPBY_LABEL`) : Mois, Bénéficiaire, Statut, Groupe, Médecin, Terminé ; toujours déplié, pas de pagination ; totaux `computeSanteTotals`.
- Fiche `renderSanteDetail` (réutilise `modal-detail`) ; `openReimbEditor`, `openPartMoiEditor` (0/50/100 % + saisie).
- Réglages `__SF_SANTE_SETTINGS` : listes + export (`exportSanteExcel`, feuille "Santé", `buildSanteExportRows` partagé avec l'export principal) / import (`rowToSanteItem`, aussi utilisé pour détecter une feuille "Santé" dans le flux principal ; réconciliation silencieuse §7.1 bis avant calcul d'id).
- Stockage IndexedDB `sante_items` (repli `sf_sante_items_v1`). Toujours dans le JSON et la synchro Drive (remplacement complet).

### 6.7 Revenus
Répartition Moi / Ma femme (`moi`/`conjoint`, `REV_PERSON_LABEL`), sans bouton flottant.
```
Store.settings.revenus = { periodFrom, periodTo, groupFilter: [], groupsExpanded,
  moi: { lines: [], transferOut: [], transferIn: [] }, conjoint: { idem } }
```
- Ligne `{id, type, label, sign, …}` : `manual` (`values {"AAAA-MM": montant}`), `recurring` (`recurCategorie` + `recurLabel` → `buildRecurringIndex`), `budgetSalary` (lit/écrit en direct `budget.salaryEntries[mois].femme`, une seule, sous Ma femme).
- `sign` ±1 appliqué dans `revLineColValue`/`revLineTotalValue` ; défaut −1 pour `transferOut`, +1 pour `transferIn`.
- Par personne : **Avant transfert** (somme de `lines`) → **Transfert** = `transferOut + transferIn` (addition pure, le signe est porté par chaque ligne) → total sur la ligne Personne.
- Miroir : `moi.transferOut` ↔ `conjoint.transferIn` (et inversement), même `id`, **signe inversé** (`revSyncTransferLineToMirror`) ; suppression des deux (`revRemoveTransferLineAndMirror`).
- Colonnes : années de toute la période (`revYearsList`), zoom mois, colonne lump pour `AAAA`. Cellules éditables manuel/salaire uniquement (`openRevValueEditor`, `openRevYearEditor` 12 mois), miroir répercuté.
- Repliage : `revCollapsedPersons`, `revCollapsedSections` (non persistés), `tr.hidden`.
- Filtres : période (`ensureRevenusPeriod`) ; Groupe (`openRevFilterSheet`, composant hiérarchique partagé), n'affecte que les lignes récurrentes (teste `groupeBase` et `getGroupeDisplay`).
- Gestion `openRevManager` ; ajout récurrent via `pickRecurringTransaction` ; transferts sans option salaire (`openAddTransferLineChoice`).
- `__SF_REVENUS.buildParamsPayloadPart()`/`applyParamsImportPart(revenus, mode)`.

---

## 7. Import
- Formats `.xlsx/.xls/.csv` et `.json` (sauvegarde v2). `findDataSheetName()` = feuille avec en-têtes Date+Montant. `__SF_IMPORT.open(file?)` : un fichier fourni est analysé directement (utilisé par l'import de l'Excel du Drive).
- Choix Transactions+Paramètres / Transactions / Paramètres, puis Ajouter/Remplacer par volet ; une feuille "Santé" est importée indépendamment de ce choix, avec le mode des Transactions.
- Dates : `raw:true` pour `.xlsx/.xls` (cellules date lues en `Date`, sans ambiguïté jour/mois) ; `raw:false` pour CSV (texte relu par `U.parseDateValue`).
- Compte et Compte cible partagent un espace de noms.
- `Store.normalizeSettings()` en fin de tout import de paramètres.

**Feuilles de paramètres** (optionnelles, `parseParametresSheet(wb)` → même forme que `buildParamsPayload()`) :

| Feuille | Colonnes | Alimente |
|---|---|---|
| Paramètres | `Groupe · Groupe de groupe · Type budget (Quotidien/Voyage) · Mois voyage · Date début (jours) · Date fin (jours) · [vide] · Compte · Type de compte · Acronyme · [vide] · Personne · Groupe de personnes · [vide] · Catégorie (classement forcé) · Type forcé (Revenus/Dépenses)` | règles groupe/personne, types/acronymes, `groupKind`/`groupVoyageMonth`/`groupDateRange`, `categoryTypeOverride` |
| Budget | `Mois · Groupe · Catégorie · Montant · Remarque · Amorti · Vue (Perso/Couple/Les deux) · Vue uniquement` | `budget.adjustments` |
| Budget_Salaires | `Mois · Moi · Ma femme` | `budget.salaryEntries` |
| Budget_Predefini | `Groupe · Catégorie · Montant mensuel` (`(Voyages)` = ligne voyages) | `budget.compare.predefined` |
| Revenus_Lignes | `Personne · Catégorie · Ligne · Type · Catégorie récurrente · Signe` | lignes Revenus (Catégorie = sentinelle `(Transfert - vers)`/`(Transfert - depuis)` ou vide) |
| Revenus_Montants | `Personne · Catégorie · Ligne · Mois · Montant` | montants des lignes manuelles |

Paramètres : "Groupe" = libellé complet ; un suffixe entre parenthèses → règle `full` (base dérivée par `U.baseGroupe`) ; Type/Mois voyage/bornes indexés sur la valeur brute. "Mois voyage" = vraie date au 1ᵉʳ du mois (`voyageMonthToDateObj`). Salaires/prédéfini : fusion mois par mois en Ajouter, remplacement en Remplacer.

### 7.1 Réconciliation interactive (Excel/CSV avec transactions existantes uniquement)
`DATA.buildReconciliation(txns)` sur Catégorie, Personne, Compte(+cible), Groupe (`U.normalize`) : correspondance → alignement silencieux ("N valeur(s) alignée(s)") ; valeurs vraiment nouvelles (variantes fusionnées) → écran `renderImportReconcile`, "Nouveau" ou "Remplacer par…".

### 7.1 bis Réconciliation silencieuse des feuilles de réglages
`U.alignToCanon(v, table)`, `U.transactionFieldCanon(champs)`, `U.valueListCanon(valeurs)` :
- Paramètres : Compte, Personne, Catégorie sur l'existant ; Groupe sur un libellé littéral exact.
- Budget : Groupe sur `mainGroup`/`extraGroups` ou libellé existant ; Catégorie via `__SF_BUDGET.categoriesForGroup`.
- Budget_Predefini : Groupe strictement sur `mainGroup`/`extraGroups` ; Catégorie idem.
- Revenus : Catégorie récurrente ; recherche de Ligne insensible casse/accents entre les deux feuilles (`findRevLine`/`findRevTransferLine`).
- Santé : Groupe + 6 listes, avant `buildSanteId`.

### 7.2 Payload `params` (`buildParamsPayload()` / `applyParamsImport(params, mode)`)
Un seul payload pour les trois portes : sauvegarde `.json`, synchro Drive, feuilles Excel.
- **Inclus** : `groupAliasRules`, `personAliasRules`, `accountTypes`, `accountAcronyms`, `accountTypeOrder`, `customBalances`, `decimals`, `groupKind`, `groupVoyageMonth`, `groupDateRange`, `categoryTypeOverride`, `budget.{adjustments, mode, scope, startMonth, endMonth, mainGroup, extraGroups, amortCategories, excludedAccounts, periodFrom, periodTo, salaryEntries, compare, mergeVoyages, mergeVoyagesGroups}`, `evolution.{view, monthBasis, merges}`, `revenus.*`.
- **JSON/Drive seulement (pas d'Excel)** : réglages budget hors ajustements/salaires/prédéfini, `compare.{enabled,mode,periodFrom,periodTo}`, `accountTypeOrder`, `decimals`, `customBalances`, `evolution.*`, `revenus.{periodFrom,periodTo,groupFilter}`.
- **Exclus** : `theme`, filtres, états dépliés (`txnGroupsExpanded`, `recurrent.groupsExpanded`, `santeGroupsExpanded`), `budget.{detailView,monthlyView,displayMode}`, `recurrent.{periodFrom,periodTo}`, `evolution.cumulative`, `hideZeroAccounts`, et **`Store.settings.drive`** (identifiants, jamais dans une sauvegarde).
- **Garde-fous** : en Remplacer, `budget.periodFrom/To` retombent sur la valeur locale si absents ; `revenus.groupFilter` adopté seulement si non vide (`[]` est truthy) ; chaque bloc n'est écrasé que si `!params.sheetsPresent || sheetsPresent.xxx` (JSON/Drive n'ont jamais `sheetsPresent` → tout est appliqué ; Excel → seulement les feuilles présentes).
- **Mode Ajouter** : dédoublonnage (`mergeRuleList`) pour règles, ajustements, `evolution.merges` ; adoption non destructive (seulement si rien de local) pour `accountTypeOrder`, `customBalances`, `excludedAccounts`, la plupart des réglages budget, le bloc `revenus` entier ; jamais touchés : `evolution.view/monthBasis`, `decimals`.

---

## 8. Export
- **Excel** (`openExportExcelScopeChoice`) : cases Transactions / Paramètres / Santé (Santé si soins présents) → un classeur ; suffixe de nom selon la sélection. Construction par `buildExcelWorkbook(scope)` (sans téléchargement, partagé avec le miroir Drive) ; `exportExcel` télécharge. Dates en vraies cellules `dd/mm/yyyy` (`U.setDateColumnFormat`, `U.isoToDateObj`, `cellDates:true`).
- **JSON** (`exportBackup`) : Transactions+Paramètres / Transactions / Paramètres ; soins Santé toujours inclus.

---

## 9. Stockage local
IndexedDB `sf_finance_db` : `transactions`, `sante_items` ; repli localStorage (`LS_KEY_LEGACY`, `sf_sante_items_v1`). `Store._persistJson` (sérialisation, IDB, repli d'urgence < 4 Mo ; seules les transactions basculent durablement en mode localStorage). `Store.settings` en localStorage `sf_settings_v1` (dont `drive`, §10). `Store.normalizeSettings()` comble les champs manquants et supprime tout reste de l'ancienne config `sync` (Gist).

---

## 10. Synchronisation Google Drive (automatique)
Relais **Google Apps Script** déployé par l'utilisateur (`google-apps-script.gs`, étapes en tête) : Application Web, **Exécuter en tant que : Moi**, **Accès : Tout le monde** ; après modification du code → Gérer les déploiements → Nouvelle version (URL `/exec` inchangée). Sinon Google renvoie une page de connexion/erreur sans CORS → l'app affiche "Connexion impossible au script Google".

**Script** : `doPost` → `handle_` entièrement protégé (toute exception renvoyée en `{ok:false, error:"Script error: …"}`), `SECRET` codé dans le script (valeur factice refusée), `LockService.tryLock(30 s)`. Actions `getJson`, `putJson` (JSON valide exigé), `getExcel`, `putExcel` (base64). `doGet` → `{"ok":true,…}` pour tester l'URL dans un navigateur. Dossier `Suivi financier` à la racine, retrouvé par la propriété de script `FOLDER_ID` (déplaçable/renommable ; renseignable à la main pour cibler un dossier existant). Mise à jour en place via Drive API v3 `uploadType=media` (même ID → historique Drive conservé), repli : fichier recréé.

**Fichiers Drive** : `suivi_financier.json` = source de vérité (`buildSyncPayload()` : transactions + `params` + `santeItems` + `santeSettings`) ; `suivi_financier.xlsx` = miroir (`buildExcelWorkbook({transactions, params, sante:true})`), **jamais relu automatiquement**.

**Côté app** (IIFE Synchronisation, `window.__SF_SYNC`) :
- `driveCall(body)` : POST `text/plain` (pas de pré-vol CORS), clé dans le corps, timeout 60 s.
- Config `Store.settings.drive = {url, key, lastSyncedAt, lastSyncedHash, lastRemoteHash, lastRemoteAt, lastExcelHash, lastExcelAt}` ; URL validée (`https://script.google.com/…/exec`) ; changer URL/clé remet les références à zéro ; Déconnecter efface la config locale seulement.
- **`smartSync`** (hash local, distant, références `lastSyncedHash`/`lastRemoteHash`) : rien sur le Drive → envoi ; identiques → rien ; seul le distant a changé → récupération (`adoptRemotePayload`, remplacement complet) ; seul le local → envoi, **sauf appareil vide face à un Drive rempli** ; appareil vide jamais synchronisé → récupération ; sinon modale de choix (avec comptes de transactions/soins). En automatique : pas de modale, état `conflict`, auto-sync suspendue jusqu'à un appui sur l'indicateur.
- **Déclencheurs** (`start()` appelé par `init()`, enrobe `saveTransactions`/`saveSanteItems`/`saveSettings`) : JSON au démarrage (+1,5 s), au retour sur l'app (≤ 1/min), sur `online`, **30 s après la dernière écriture** (`AUTO_DELAY_MS`), immédiatement au passage en arrière-plan (`visibilitychange`/`pagehide`) s'il y a un envoi en attente. Écriture sans changement de hash → aucun appel réseau.
- **Excel** (`maybeUploadExcel`) : au passage en arrière-plan et au démarrage (rattrapage), si données changées depuis `lastExcelHash` **et** ≥ 1 h depuis `lastExcelAt` ; premier dépôt immédiat après le premier envoi ; pas de redépôt après une récupération. Erreur Excel affichée dans la feuille seulement.
- **Feuille** : état, Synchroniser maintenant, Forcer récupération (confirmation) / sauvegarde, bloc Excel (dernier dépôt, À jour / en attente), **Importer l'Excel** (→ `__SF_IMPORT.open(file)`) et **Mettre à jour l'Excel** (immédiat). Rendus d'arrière-plan ignorés si un champ de saisie a le focus.
- **Indicateur** `#btn-sync` (`computeStatus`) : unconfigured / syncing / synced / unsynced (envoi en attente) / error / conflict.
- API : `{open, updateIndicator, start, autoSync, smartSync, pushToDrive, pullFromDrive, pushExcelNow, importExcelFromDrive, maybeUploadExcel}`.

---

## 11. Interface (One UI)
Palette bleue, cartes arrondies, thème clair/sombre auto ou forcé. Version affichée sous le titre (`APP_VERSION` sans `v `). Données d'exemple depuis les écrans vides. Réinitialisation complète dans le Menu ☰ (efface aussi les soins).
- `.chev-rot`, `.pill-toggle`/`.pill-toggle-wrap`, `.txn-tap-row` (ligne tactile réutilisable).
- `position:sticky` échoue silencieusement dans une grille CSS ou une `.card` → utiliser `<table>` pour les colonnes figées ; un en-tête sticky doit être enfant direct du conteneur à couvrir.
- `.evo-table tr.hidden{display:none}` pour les lignes repliables ; `.evo-col-label` en `white-space:normal`.

---

## 12. Architecture (IIFE, `window.__SF_*`, ordre d'exécution)

| Module | Expose | Contenu clé |
|---|---|---|
| Utils | `__SF_UTILS` | formatage/dates, icônes, `baseGroupe`, `baseRemarque`, `computeTransform`, `enableDragScroll`, décimales / 0€, `capitalize`/`fmtMonthYearCap`, `fmtWholeSigned`/`signClass`, `isoToDateObj`/`setDateColumnFormat`, `hierarchyCheckListHtml`/`wireHierarchyCheckList`, `alignToCanon`/`transactionFieldCanon`/`valueListCanon`, `xlsxAvailable`/`XLSX_MISSING_MSG` |
| Store + Import/Parsing | `__SF_STORE`, `__SF_DATA` | persistance, `normalizeSettings`, `parseFile`, `importTxns`, `rowToTxn`, `buildTxnId`, `generateDemoData`, `buildReconciliation`, `parseParametresSheet` |
| Filtres + Agrégations | `__SF_FILTERS`, `__SF_AGG` | filtres ; résolution groupe/personne/compte ; `getCalcMonthYM` ; hiérarchies (`groupeHierarchyChrono`/`Full`, `personneHierarchy`) ; dates de groupe ; `labelResolvesToSuperGroupe` ; moteurs Évolution et Récurrent |
| Contrôleur | `__SF_APP` | shell, onglets (`MORE_TABS`), thème, état partagé, `openOverlay`/`closeOverlay`, `confirmAction`, `refreshCurrentTab`, `showToast` |
| Vues | `__SF_VIEWS.*`, `__SF_BUDGET`, `__SF_REVENUS` | Comptes, Transactions, Budget, Évolution, Récurrent, Revenus (§6) |
| Feuille de filtre | `__SF_FILTERSHEET` | menu 2 étapes, `NESTED_DIM_CONFIG` |
| Import/Menu | `__SF_IMPORT`, `__SF_MENU` | `open(file?)`, `applyParamsImport` ; `open`, `buildParamsPayload`, `buildExcelWorkbook` ; export, réinitialisation |
| Synchronisation | `__SF_SYNC` | §10 |
| Santé | `__SF_SANTE`, `__SF_SANTE_DATA`, `__SF_SANTE_FILTERS`, `__SF_SANTE_FILTERSHEET`, `__SF_SANTE_SETTINGS` | §6.6 |
| `createAliasManager(cfg)` | `__SF_GROUPMGR`, `__SF_PERSONMGR`, `__SF_ACCTTYPE` | §2.4, §3 |
| PWA | — | enregistrement du SW, bannière de mise à jour (§14) |

**Extensions** :
- Nouvelle dimension type Groupe/Personne : `getXDisplay`/`findXAliasRule`/`distinctBaseX`/`xHierarchy` + `createAliasManager` + entrée `NESTED_DIM_CONFIG`.
- Nouvelle dimension Évolution : `buildEvolutionIndex` **et** `txnMatchesDimTarget`, cohérents.
- Picker : `renderPickerList` + `opts.hierarchy = {sections, flat, preSorted?, selectableHeaders?, expandedByDefault?, kindFn?}`.
- Nouvel onglet secondaire : `MORE_TABS`/`MORE_TAB_KEYS`, `view-*`, `showFloat`, `refreshCurrentTab` et dispatch des boutons (Santé = modèle).
- Toujours réutiliser les helpers partagés (formateurs, liste cochable hiérarchique, `openBudgetKindSheet`, `Store._persistJson`, `driveCall`/`runSyncOp`) plutôt que d'en recopier une variante.

---

## 13. Limites connues
- Conflits de synchro résolus au hash global (choix d'une version entière, pas de fusion). Modifications faites dans l'Excel du Drive prises en compte seulement via "Importer l'Excel".
- Mapping d'import par nom d'en-tête normalisé, jamais par position.
- Règles de groupe sans fenêtre de dates ; `groupDateRange` déclaratif (pas de contrôle de cohérence) ; "aujourd'hui" seulement pour `dateTo` ; suffixe de groupe non reconnu → repli silencieux sur la plus ancienne transaction.
- Fusions Évolution indépendantes des règles groupe/personne ; création seulement via JSON/synchro.
- `evolution.cumulative` et `hideZeroAccounts` locaux à l'appareil.
- Réconciliation interactive limitée à Catégorie/Personne/Compte/Groupe des transactions ; jamais pour un `.json`.
- Santé : l'import Ajouter ignore un soin dont l'id existe (pas de mise à jour) ; comptages du filtre Groupe issus des transactions ; pas de taux de change. Feuille Santé du flux principal = mode des Transactions.
- Revenus : lien récurrent par `(recurCategorie, recurLabel)`, cassé silencieusement si la catégorie ou la remarque change ; filtre Groupe sans effet sur lignes manuelles/salaire.
- Navigation fiche transaction : liste figée pendant l'ouverture ; absente hors liste principale.
- Ajustements : repli groupe précis → groupe de groupe uniquement.
- Sans SheetJS (hors ligne au premier chargement), seules les fonctions Excel sont indisponibles (`U.XLSX_MISSING_MSG`).

---

## 14. PWA
- `manifest.json` : icônes standard + maskable, `display:standalone`, couleurs alignées sur `--bg`.
- `service-worker.js` : cache `sf-shell-${SW_VERSION}` (**`SW_VERSION` = `APP_VERSION`** à chaque livraison) ; `index.html` réseau d'abord ; autres fichiers same-origin cache d'abord ; CDN en stale-while-revalidate (`sf-runtime`, réponses exploitables seulement) ; requêtes non-GET jamais interceptées ; anciens caches supprimés à l'activation.
- Bannière "Nouvelle version disponible" → `SKIP_WAITING` puis rechargement au `controllerchange`, **seulement si `refreshRequested`** : `clients.claim()` déclenche aussi `controllerchange` à la première installation.
