INSERT INTO `client` (`numCli`, `nomCli`, `precisionCli`, `villeCli`) VALUES
(1, 'CAVANA', 'Marie', 'LA ROCHELLE'),
(2, 'BURLET', 'Michel', 'LAGORD'),
(3, 'PEUTOT', 'Maurice', 'LAGORD'),
(4, 'ORGEVAL', 'Centrale d achats', 'SURGERES');

INSERT INTO `concerner` (`numPdt`, `numSort`, `qteSort`) VALUES
(20241, 1, 300),
(20241, 2, 400),
(20242, 1, 200),
(20243, 1, 100),
(20243, 2, 500);

INSERT INTO `coûte` (`numPdt`, `Annee`, `Prix_Achat`, `Prix_Vente`) VALUES
(1, 2023, 270, 280),
(1, 2024, 270, 290),
(1, 2025, 240, 300),
(2, 2023, 3900, 9500),
(2, 2024, 3800, 10000),
(2, 2025, 3500, 9000);

INSERT INTO `entree` (`numEnt`, `dateEnt`, `qteEnt`, `numPdt`, `numSau`) VALUES
(20241, '2024-06-16 00:00:00', 1000, 1, 1),
(20242, '2024-06-18 00:00:00', 500, 2, 1),
(20243, '2024-07-10 00:00:00', 1500, 2, 2);

INSERT INTO `produit` (`numPdt`, `stockPdt_t_`, `libPdt`) VALUES
(1, 2000, 'Gros sel'),
(2, 1000, 'Fleur de sel');

INSERT INTO `saunier` (`numSau`, `nomSau`, `prenomSau`, `villeSau`) VALUES
(1, 'YVAN', 'Pierre', 'Ars-en-Ré'),
(2, 'PETIT', 'Marc', 'Loix');

INSERT INTO `sortie` (`numSort`, `datSort`, `numCli`) VALUES
(20241, '2024-07-16 00:00:00', 1),
(20242, '2024-07-18 00:00:00', 1),
(20243, '2024-08-10 00:00:00', 2);