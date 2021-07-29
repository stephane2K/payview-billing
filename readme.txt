Factupayiew.py
le 29.07.2021
installation de la facturation Payview + conversion au format de commande SAP (.csv)


1/
récupérer les sources :
sur https://github.com/stephane2K/payview-billing
cliquer sur 
CODE 
Download Zip

2/ dézipper cela dans un répertoire nommé plus tard "répertoire d'exécution" puis mettre à jour les variables :


FactuPayview.py:ligne 25 config_path  ==> répertoire contenant parametres.ini (envoyé mar. 9.25 par mail).

Mettre à jour le fichier paramètres.ini avec :
[ EMAIL MOT DE PASSE PAYVIEW ADMIN ]
PAYVIEW_ADMIN_EMAIL                  ==> login
PAYVIEW_ADMIN_MDP                    ==> password

[CHEMINS DES FICHIERS D'ENTREES] :
CLIENTS_IGNORES                      ==> clients_non_facturés.txt
FICHIER_SIM_PRE                      ==> SIMs_de_pret.xlsx
FICHIER_CORRESPONDANCE_NOMS_CLIENTS  ==> correspPayViewPassPort.xlsx
[CHEMINS DES FICHIERS GENERES]
DOSSIER_GENERATION_RESULTATS         ==> Generated (répertoire de génération)

(possibilité d'utiliser des liens réseau r"//frpfil/...")


dans commandeSAP.py

ligne 13           ==> racine= répertoire Generated choisi ci-dessus
ligne 22           ==>  nom_fichier= répertoire Generated/nom de la synthèse ADV
ligne 30           ==>  path_lut : chemin et nom du fichier LUT.xslx
ligne 32           ==>  path_concat : chemin et nom du fichier .csv généré à la fin.
les autres variables sont des constantes.

3/ Mettre à jour le login / password Passport.

dans mshAPI.py, ligne 74, mettre un login et mot de passe valable
(celui entré en exemple est KO)

4/ exécuter le script
dans le répertoire d'exécution, par exemple en commandes DOS, taper :
# pyhon Factupayview.py

5/ le résultat est généré dans le fichier .csv, prêt à être uploadé dans SAP
sinon, l'ADV peut traiter manuellement synthèse ADV.xslx
les deux fichiers leur sont utiles pour faire des vérifications.

----------------------------------------------------------------------------------
mise à jour du 29.07.2021.

commanades SAP.py

PAS_SIM_OVERFEE non pris en compte pour certains clients.







