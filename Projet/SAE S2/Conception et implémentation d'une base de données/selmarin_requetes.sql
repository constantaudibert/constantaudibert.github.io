--les 3 meilleurs contributeurs de la coopérative
SELECT S.numSau, S.nomSau, S.prenomSau, SUM(E.qteEnt) AS Quantite_Totale_Fournie
FROM SAUNIER S
JOIN ENTREE E ON S.numSau = E.numSau
GROUP BY S.numSau, S.nomSau, S.prenomSau
ORDER BY Quantite_Totale_Fournie DESC
LIMIT 3;

--meilleurs clients
SELECT C.numCli, C.nomCli, SUM(CO.qteSort * P.prixVente) AS Chiffre_Affaires
FROM CLIENT C
JOIN SORTIE S ON C.numCli = S.numCli
JOIN CONCERNER CO ON S.numSort = CO.numSort
JOIN PRODUIT P ON CO.numPdt = P.numPdt
GROUP BY C.numCli, C.nomCli, C.villeCli
ORDER BY Chiffre_Affaires DESC
LIMIT 5;

--insertion Saunier garnier
INSERT INTO SAUNIER VALUES (3,"GARNIER","François","La Flotte")

--état des stocks 
SELECT numPdt, libPdt, stockPdt_t_ 
FROM PRODUIT;

--état des stocks pour chaque saunier
select  nomSau, prenomSau, E.NumPdt, libPdt stockPdt_t_
FROM S SAUNIER, E ENTREE,  P PRODUIT
where S.numSau=E.numSau
and E.numPdt=P.numPdt

--évolution du prix de ventes au fil des années
SELECT A.Annee, C.Prix_Vente, P.libPdt FROM PRODUIT P
JOIN Coûte C ON P.numPdt = C.numPdt
JOIN AnneePrix A ON C.Annee = A.Annee
ORDER BY P.libPdt,A.Annee

-- Création d'un utilisateur "garnier" avec un mot de passe
CREATE USER 'garnier'@'localhost' IDENTIFIED BY '1234';

-- Création de la vue "CA_Annuel"
CREATE VIEW CA_Annuel
SELECT YEAR(S.datSort) AS Annee,SUM(CO.qteSort * P.prixVente) AS Chiffre_Affaires
FROM SORTIE S
JOIN CONCERNER CO ON S.numSort = CO.numSort
JOIN PRODUIT P ON CO.numPdt = P.numPdt
GROUP BY Annee
ORDER BY Annee DESC;

-- Donner les droits de SELECT à l'utilisateur "garnier" sur la vue "CA_Annuel"
GRANT SELECT ON CA_Annuel TO 'garnier'@'localhost';

-- Utilisation de la vue CA_Annuel
SELECT * FROM CA_Annuel;

--Marge moyenne par prix de ventes
SELECT P.numPdt, P.libPdt, AVG(C.Prix_Vente - C.Prix_Achat) AS Marge_Moyenne
FROM PRODUIT P
JOIN COÛTE C ON P.numPdt = C.numPdt
GROUP BY P.numPdt, P.libPdt
ORDER BY Marge_Moyenne DESC;

--Produit jamais vendus
SELECT DISTINCT P.numPdt, P.libPdt
FROM PRODUIT P
WHERE P.numPdt NOT IN (SELECT DISTINCT C.numPdt FROM CONCERNER C);

