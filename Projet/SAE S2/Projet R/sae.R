library(ggplot2)
library(maps)

# Chargement et filtrage
datasup2023 <- read.csv2(file ="fr-esr-parcoursup_2023.csv")
datasup2023 <- datasup2023[datasup2023$Code.départemental.de.l.établissement == "37", ]

datasup2024 <- read.csv2(file = "fr-esr-parcoursup_2024.csv")
datasup2024 <- datasup2024[datasup2024$Code.départemental.de.l.établissement == "37", ]

# Graphique 1 : Histogramme des capacités
ggplot(datasup2024, aes(x = as.numeric(Capacité.de.l.établissement.par.formation))) +
  geom_histogram(binwidth = 10, fill = "#3182bd", color = "black") +
  labs(
    title = "Distribution des capacités d’accueil par formation",
    x = "Capacité d’accueil",
    y = "Nombre de formations"
  ) +
  theme_minimal()

# Graphique 2 : Mentions au bac
mentions_cols <- c(
  "Dont.effectif.des.admis.néo.bacheliers.sans.mention.au.bac",
  "Dont.effectif.des.admis.néo.bacheliers.avec.mention.Assez.Bien.au.bac",
  "Dont.effectif.des.admis.néo.bacheliers.avec.mention.Bien.au.bac",
  "Dont.effectif.des.admis.néo.bacheliers.avec.mention.Très.Bien.au.bac",
  "Dont.effectif.des.admis.néo.bacheliers.avec.mention.Très.Bien.avec.félicitations.au.bac"
)
mention_names <- c("sans", "assez_bien", "bien", "tres_bien", "felicitations")
mention_vals <- colSums(sapply(datasup2024[mentions_cols], function(x) as.numeric(gsub(",", ".", x))), na.rm = TRUE)

ggplot(data.frame(Mention = factor(mention_names, levels = mention_names), Nombre = mention_vals),
       aes(x = Mention, y = Nombre, fill = Mention)) +
  geom_col(show.legend = FALSE) +
  scale_fill_manual(values = c(
    "sans" = "#deebf7", "assez_bien" = "#9ecae1", "bien" = "#6baed6",
    "tres_bien" = "#3182bd", "felicitations" = "#08519c"
  )) +
  labs(
    title = "Distribution des mentions au baccalauréat parmi les admis récents",
    x = "Mention obtenue au bac",
    y = "Nombre d'admis"
  ) +
  theme_minimal()

# Graphique 3 : Taux d’accès moyen par filière
datasup2024$Taux.d.accès <- as.numeric(gsub(",", ".", datasup2024$Taux.d.accès))
moy <- tapply(datasup2024$Taux.d.accès, datasup2024$Filière.de.formation.très.agrégée, mean, na.rm = TRUE)

ggplot(data.frame(Filière = names(moy), Taux_moyen = moy),
       aes(x = Filière, y = Taux_moyen, fill = Taux_moyen)) +
  geom_col() +
  scale_fill_gradient(low = "#deebf7", high = "#08306b") +
  labs(
    title = "Taux d'accès moyen par filière de formation",
    x = "Filière de formation",
    y = "Taux d'accès moyen (%)"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Graphique 4 : Répartition des 5 filières principales par commune (top 3)
top_communes <- names(sort(table(datasup2024$Commune.de.l.établissement), decreasing = TRUE))[1:3]
top_filieres <- names(sort(table(datasup2024$Filière.de.formation.très.agrégée), decreasing = TRUE))[1:5]
df_sub <- datasup2024[datasup2024$Commune.de.l.établissement %in% top_communes &
                        datasup2024$Filière.de.formation.très.agrégée %in% top_filieres, ]
tab <- table(df_sub$Commune.de.l.établissement, df_sub$Filière.de.formation.très.agrégée)
df_tab <- as.data.frame(tab)
names(df_tab) <- c("Commune", "Filière", "n")

ggplot(df_tab, aes(x = Commune, y = n, fill = Filière)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  scale_fill_manual(values = c("#deebf7", "#9ecae1", "#6baed6", "#3182bd", "#08519c", "#bdd7e7")) +
  labs(
    title = "Répartition des 5 principales filières par commune (top 3) - Département 37",
    x = "Commune",
    y = "Nombre d’établissements",
    fill = "Filière"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Graphique 5 : Effectifs 2023 vs 2024 par statut
statut_2023 <- ifelse(datasup2023$Statut.de.l.établissement.de.la.filière.de.formation..public..privé.. == "Public", "public", "privé")
statut_2024 <- ifelse(datasup2024$Statut.de.l.établissement.de.la.filière.de.formation..public..privé.. == "Public", "public", "privé")

df_eff <- data.frame(
  année = c(rep("2023", nrow(datasup2023)), rep("2024", nrow(datasup2024))),
  statut_simple = c(statut_2023, statut_2024),
  Effectif = c(as.numeric(datasup2023$Effectif.total.des.candidats.pour.une.formation),
               as.numeric(datasup2024$Effectif.total.des.candidats.pour.une.formation))
)

ggplot(df_eff, aes(x = année, y = Effectif, color = statut_simple)) +
  geom_jitter(width = 0.2, alpha = 0.5, size = 2) +
  scale_color_manual(values = c("public" = "#08306b", "privé" = "#deebf7")) +
  labs(
    title = "Comparaison des effectifs par statut d’établissement (2023 vs 2024)",
    subtitle = "Chaque point = une formation ; losange noir = moyenne",
    x = "Année",
    y = "Effectif total des candidats",
    color = "Statut"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Graphique 6 : Carte géographique
coords <- strsplit(as.character(datasup2024$Coordonnées.GPS.de.la.formation), ",")
latlon <- do.call(rbind, lapply(coords, function(x) if (length(x) == 2) as.numeric(x) else c(NA, NA)))
datasup2024$lat <- latlon[, 1]
datasup2024$lon <- latlon[, 2]

ggplot() +
  geom_polygon(data = map_data("france"), aes(x = long, y = lat, group = group),
               fill = "gray90", color = "white") +
  geom_point(data = datasup2024[!is.na(datasup2024$lat), ],
             aes(x = lon, y = lat),
             color = "#3182bd", size = 2, alpha = 0.7) +
  coord_fixed(xlim = c(0.4, 1.6), ylim = c(46.7, 47.7)) +
  labs(
    title = "Localisation des formations Parcoursup - Indre-et-Loire (37)",
    x = "Longitude", y = "Latitude"
  ) +
  theme_minimal()









# Sommes totales par année et statut
total_2023_public <- sum(datasup2023$Effectif.total.des.candidats.pour.une.formation[statut_2023 == "public"], na.rm = TRUE)
total_2023_prive <- sum(datasup2023$Effectif.total.des.candidats.pour.une.formation[statut_2023 == "privé"], na.rm = TRUE)
total_2024_public <- sum(datasup2024$Effectif.total.des.candidats.pour.une.formation[statut_2024 == "public"], na.rm = TRUE)
total_2024_prive <- sum(datasup2024$Effectif.total.des.candidats.pour.une.formation[statut_2024 == "privé"], na.rm = TRUE)

# Création du dataframe
df_totals <- data.frame(
  année = rep(c("2023", "2024"), each = 2),
  statut = rep(c("public", "privé"), times = 2),
  total_effectif = c(total_2023_public, total_2023_prive,
                     total_2024_public, total_2024_prive)
)

# Calcul des pourcentages par année
pct_2023 <- df_totals$total_effectif[df_totals$année == "2023"] / sum(df_totals$total_effectif[df_totals$année == "2023"]) * 100
pct_2024 <- df_totals$total_effectif[df_totals$année == "2024"] / sum(df_totals$total_effectif[df_totals$année == "2024"]) * 100

df_totals$pourcentage <- c(pct_2023, pct_2024)

ggplot(df_totals, aes(x = année, y = pourcentage, fill = statut)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values = c("public" = "#08306b", "privé" = "#deebf7")) +
  labs(
    title = "Répartition des effectifs totaux par statut et année",
    x = "Année",
    y = "Pourcentage des effectifs",
    fill = "Statut"
  ) +
  theme_minimal()

