-- phpMyAdmin SQL Dump
-- version 4.7.0
-- https://www.phpmyadmin.net/
--
-- Hôte : 127.0.0.1
-- Généré le :  lun. 24 mars 2025 à 10:56
-- Version du serveur :  5.7.17
-- Version de PHP :  5.6.30

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de données :  `fabianconstant`
--

-- --------------------------------------------------------

--
-- Structure de la table `anneeprix`
--

CREATE TABLE `anneeprix` (
  `Annee` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `anneeprix`
--

INSERT INTO `anneeprix` (`Annee`) VALUES
(2023),
(2024);

-- --------------------------------------------------------

--
-- Structure de la table `client`
--

CREATE TABLE `client` (
  `numCli` int(11) NOT NULL,
  `nomCli` varchar(50) NOT NULL,
  `precisionCli` varchar(50) NOT NULL,
  `villeCli` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `client`
--

INSERT INTO `client` (`numCli`, `nomCli`, `precisionCli`, `villeCli`) VALUES
(1, 'CAVANA', 'Marie', 'LA ROCHELLE'),
(2, 'BURLET', 'Michel', 'LAGORD'),
(3, 'PEUTOT', 'Maurice', 'LAGORD'),
(4, 'ORGEVAL', 'Centrale d\'achats', 'SURGERES');

-- --------------------------------------------------------

--
-- Structure de la table `concerner`
--

CREATE TABLE `concerner` (
  `numPdt` int(11) NOT NULL,
  `numSort` int(11) NOT NULL,
  `qteSort` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `concerner`
--

INSERT INTO `concerner` (`numPdt`, `numSort`, `qteSort`) VALUES
(20241, 1, 300),
(20241, 2, 400),
(20242, 1, 200),
(20243, 1, 100),
(20243, 2, 500);

-- --------------------------------------------------------

--
-- Structure de la table `coûte`
--

CREATE TABLE `coûte` (
  `numPdt` int(11) NOT NULL,
  `Annee` int(11) NOT NULL,
  `Prix_Achat` int(11) DEFAULT NULL,
  `Prix_Vente` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `coûte`
--

INSERT INTO `coûte` (`numPdt`, `Annee`, `Prix_Achat`, `Prix_Vente`) VALUES
(1, 2023, 270, 280),
(1, 2024, 270, 290),
(1, 2025, 240, 300),
(2, 2023, 3900, 9500),
(2, 2024, 3800, 10000),
(2, 2025, 3500, 9000);

-- --------------------------------------------------------

--
-- Structure de la table `entree`
--

CREATE TABLE `entree` (
  `numEnt` int(11) NOT NULL,
  `dateEnt` datetime NOT NULL,
  `qteEnt` int(11) NOT NULL,
  `numPdt` int(11) NOT NULL,
  `numSau` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `entree`
--

INSERT INTO `entree` (`numEnt`, `dateEnt`, `qteEnt`, `numPdt`, `numSau`) VALUES
(20241, '2024-06-16 00:00:00', 1000, 1, 1),
(20242, '2024-06-18 00:00:00', 500, 2, 1),
(20243, '2024-07-10 00:00:00', 1500, 2, 2);

-- --------------------------------------------------------

--
-- Structure de la table `produit`
--

CREATE TABLE `produit` (
  `numPdt` int(11) NOT NULL,
  `stockPdt_t_` int(11) NOT NULL,
  `libPdt` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `produit`
--

INSERT INTO `produit` (`numPdt`, `stockPdt_t_`, `libPdt`) VALUES
(1, 2000, 'Gros sel'),
(2, 1000, 'Fleur de sel');

-- --------------------------------------------------------

--
-- Structure de la table `saunier`
--

CREATE TABLE `saunier` (
  `numSau` int(11) NOT NULL,
  `nomSau` varchar(50) NOT NULL,
  `prenomSau` varchar(50) NOT NULL,
  `villeSau` varchar(50) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `saunier`
--

INSERT INTO `saunier` (`numSau`, `nomSau`, `prenomSau`, `villeSau`) VALUES
(1, 'YVAN', 'Pierre', 'Ars-en-Ré'),
(2, 'PETIT', 'Marc', 'Loix');

-- --------------------------------------------------------

--
-- Structure de la table `sortie`
--

CREATE TABLE `sortie` (
  `numSort` int(11) NOT NULL,
  `datSort` datetime NOT NULL,
  `numCli` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Déchargement des données de la table `sortie`
--

INSERT INTO `sortie` (`numSort`, `datSort`, `numCli`) VALUES
(20241, '2024-07-16 00:00:00', 1),
(20242, '2024-07-18 00:00:00', 1),
(20243, '2024-08-10 00:00:00', 2);

--
-- Index pour les tables déchargées
--

--
-- Index pour la table `anneeprix`
--
ALTER TABLE `anneeprix`
  ADD PRIMARY KEY (`Annee`);

--
-- Index pour la table `client`
--
ALTER TABLE `client`
  ADD PRIMARY KEY (`numCli`);

--
-- Index pour la table `concerner`
--
ALTER TABLE `concerner`
  ADD PRIMARY KEY (`numPdt`,`numSort`),
  ADD KEY `numSort` (`numSort`);

--
-- Index pour la table `coûte`
--
ALTER TABLE `coûte`
  ADD PRIMARY KEY (`numPdt`,`Annee`),
  ADD KEY `Annee` (`Annee`);

--
-- Index pour la table `entree`
--
ALTER TABLE `entree`
  ADD PRIMARY KEY (`numEnt`),
  ADD KEY `numPdt` (`numPdt`),
  ADD KEY `numSau` (`numSau`);

--
-- Index pour la table `produit`
--
ALTER TABLE `produit`
  ADD PRIMARY KEY (`numPdt`);

--
-- Index pour la table `saunier`
--
ALTER TABLE `saunier`
  ADD PRIMARY KEY (`numSau`);

--
-- Index pour la table `sortie`
--
ALTER TABLE `sortie`
  ADD PRIMARY KEY (`numSort`),
  ADD KEY `numCli` (`numCli`);

--
-- Contraintes pour les tables déchargées
--

--
-- Contraintes pour la table `concerner`
--
ALTER TABLE `concerner`
  ADD CONSTRAINT `concerner_ibfk_1` FOREIGN KEY (`numPdt`) REFERENCES `produit` (`numPdt`),
  ADD CONSTRAINT `concerner_ibfk_2` FOREIGN KEY (`numSort`) REFERENCES `sortie` (`numSort`);

--
-- Contraintes pour la table `coûte`
--
ALTER TABLE `coûte`
  ADD CONSTRAINT `coûte_ibfk_1` FOREIGN KEY (`numPdt`) REFERENCES `produit` (`numPdt`),
  ADD CONSTRAINT `coûte_ibfk_2` FOREIGN KEY (`Annee`) REFERENCES `anneeprix` (`Annee`);

--
-- Contraintes pour la table `entree`
--
ALTER TABLE `entree`
  ADD CONSTRAINT `entree_ibfk_1` FOREIGN KEY (`numPdt`) REFERENCES `produit` (`numPdt`),
  ADD CONSTRAINT `entree_ibfk_2` FOREIGN KEY (`numSau`) REFERENCES `saunier` (`numSau`);

--
-- Contraintes pour la table `sortie`
--
ALTER TABLE `sortie`
  ADD CONSTRAINT `sortie_ibfk_1` FOREIGN KEY (`numCli`) REFERENCES `client` (`numCli`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
