library(ggplot2)
library(dplyr)
library(sf)
datasup2023 <- read.csv2(file = "fr-esr-parcoursup_2023.csv")
datasup2024 <- read.csv2(file = "fr-esr-parcoursup_2024.csv")


#names(datasup2023)<- c("Session",
datasup2023
datasup372023 <- datasup2023[ substr(datasup2023$Code.UAI.de.l.établissement, 1, 3) == "037", ]


# ── 2. Préparer les données : effectifs par statut ─────────────────────────────
# Remplace "Statut.de.l.établissement..." par le nom exact de ta colonne
df_statut <- as.data.frame(table(datasup372023$Statut.de.l.établissement.de.la.filière.de.formation..public..privé.))
names(df_statut) <- c("Statut", "Effectif")

# ── 3. Calculer la part en % (facultatif mais pratique pour les labels) ────────
df_statut$Pourcent <- round(df_statut$Effectif / sum(df_statut$Effectif) * 100, 1)

# ── 4. Tracer le pie chart avec ggplot2 ────────────────────────────────────────
ggplot(df_statut, aes(x = "", y = Effectif, fill = Statut)) +
  geom_bar(width = 1, stat = "identity") +          # "anneau" vertical
  coord_polar(theta = "y") +                        # conversion en cercle
  geom_text(aes(label = paste0(Pourcent, "%")),
            position = position_stack(vjust = 0.5), # place les % au centre
            size = 4) +
  labs(title = "Répartition public / privé – Parcoursup 2024 (dépt. 37)",
       fill  = "Statut") +
  theme_void() +                                    # enlève axes et fond
  theme(plot.title = element_text(hjust = 0.5))     # centre le titre


# Supposons que ta colonne Coordonnées.GPS.de.la.formation soit sous forme "lat,long" ou "long,lat"
# Il faut d'abord extraire les coordonnées
datasup372023 <- datasup372023 %>%
  tidyr::separate(Coordonnées.GPS.de.la.formation, into = c("lat", "lon"), sep = ",", convert = TRUE)



datasup_sf <- st_as_sf(datasup372023, coords = c("lon", "lat"), crs = 4326)

# Charger la carte du département (par exemple avec le package rnaturalearth)
library(rnaturalearth)
library(rnaturalearthdata)
library(rgeos)

france <- ne_states(country = "France", returnclass = "sf")
dep_carte <- france %>% filter(gns_name == "Indre-et-Loire") # adapte selon le nom exact

# Plot avec taille selon effectif total des candidats
ggplot() +
  geom_sf(data = dep_carte, fill = "white", color = "black") +
  geom_sf(data = datasup_sf, aes(size = Effectif.total.des.candidats.pour.une.formation, color = Taux.d.accès), alpha = 0.7) +
  scale_color_viridis_c(option = "plasma") +
  labs(title = "Carte des formations dans le département 79",
       size = "Effectif total candidats",
       color = "Taux d'accès") +
  theme_minimal()


