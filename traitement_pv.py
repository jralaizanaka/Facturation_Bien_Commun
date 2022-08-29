import datetime as dt
import numpy as np
import os
import streamlit
import xlwings as xw




QUEL_MOIS = {1: "janvier", 2: "février", 3: "mars", 4: "avril",
             5: "mai", 6: "juin", 7: "juillet", 8: "août",
             9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"}


class PV:
    """ Create minutes. """

    def __init__(self,
                 ref_client,
                 # Customer reference.
                 nom_client,
                 # Customer name.
                 num_commande,
                 # Order number.
                 ref_projet,
                 # Project reference.
                 num_facture,
                 # Bill number.
                 tjm,
                 # Average daily rate (Taux journalier moyen).
                 desc_fact,
                 # Invoice description.
                 destinataire,
                 # Email address of the recipient.
                 ref_fournisseur,
                 # Supplier reference.
                 num_lot,
                 # Batch number.
                 nb_jours_commandés,
                 # Number of days ordered (prévues).
                 nb_jours_facturés,
                 # Number of days billed (réalisés)
                 mois_de_facturation,
                 # Billing month.
                 intitulé_remise="",
                 # Discount title.
                 prix_unitaire_remise="",
                 # Discount price.
                 nb_unité_remise="",
                 # Discount unit number.
                 montant_remise="",
                 # Discount amount.
                 intitulé_frais_annexe="",
                 # Additional costs title.
                 prix_unitaire_frais_annexe="",
                 # Additional costs price.
                 nb_unité_frais_annexe="",
                 #Additionnal cost unit number.
                 montant_frais_annexe="",
                 # Additional costs amount.
                 sujet_email="",
                 # Email subject.
                 contenu_email=""
                 # Email content

                 ):
        self.ref_client = ref_client
        self.nom_client = nom_client
        self.num_commande = num_commande
        self.ref_projet = ref_projet
        self.num_facture = num_facture
        self.tjm = tjm
        self.desc_fact = desc_fact
        self.destinataire = destinataire
        self.ref_fournisseur = ref_fournisseur
        self.num_lot= num_lot
        self.ref_fournisseur = ref_fournisseur
        self.nb_jours_commandés = nb_jours_commandés
        self.nb_jours_facturés = nb_jours_facturés
        self.mois_de_facturation = mois_de_facturation
        self.intitulé_remise = intitulé_remise
        self.prix_unitaire_remise = prix_unitaire_remise
        self.nb_unité_remise = nb_unité_remise
        self.montant_remise = montant_remise
        self.intitulé_frais_annexe = intitulé_frais_annexe
        self.prix_unitaire_frais_annexe = prix_unitaire_frais_annexe
        self.nb_unité_frais_annexe = nb_unité_frais_annexe
        self.montant_frais_annexe = montant_frais_annexe
        self.sujet_email = sujet_email
        self.contenu_email = contenu_email


    def rempli_pv(self,template):
        """ Open the empty template file "template" and fill it. """

        wk = xw.books.open(template)
        sheet = wk.sheets("génération_pv")
        # Select "generation_pv" sheet.

        sheet.range("C5").value = f"{QUEL_MOIS[self.mois_de_facturation.month]} {self.mois_de_facturation.year} "
        # Month and year taken into account.
        sheet.range("c9").value = f" Bien Commun  {self.ref_fournisseur}"
        # Supplier reference.
        sheet.range("e9").value = f"N° commande {self.nom_client} :"
        # Assignment of the customer's name to the order number.
        sheet.range("g9").value = f"AFR {int(self.num_commande)}"
        # Order Number.
        sheet.range("g11").value = self.ref_projet
        # Project reference.
        sheet.range("f14").value = f" {self.num_lot} / Prestation de {QUEL_MOIS[self.mois_de_facturation.month]} {self.mois_de_facturation.year} "
        # Number and title batch.
        sheet.range("b20").value = self.desc_fact
        # Bill description.
        sheet.range("c20").value = self.nb_jours_facturés
        # Number of billed days.
        sheet.range("f20").value = float(self.nb_jours_commandés)*float(self.tjm)
        # Expected amount.
        sheet.range("g20").value = float(self.nb_jours_facturés)*float(self.tjm)
        # Amount charged.

        if self.montant_remise != 0:
            sheet.range("b21").value = self.intitulé_remise
            sheet.range("c21").value = self.nb_unité_remise
            sheet.range("g21").value = self.montant_remise
        if self.montant_frais_annexe != 0:
            sheet.range("b22").value = self.intitulé_frais_annexe
            sheet.range("c22").value = self.nb_unité_frais_annexe
            sheet.range("g22").value = self.montant_frais_annexe

    def convert_to_pdf(self, enregistrer_sous, template):
        """ Print in PDF """
        wk = xw.Book(template)
        sheet = wk.sheets("génération_pv")
        nom_du_fichier = f"PVàSigner_{self.ref_client}_facture{self.num_facture}_" \
                         f"{QUEL_MOIS[self.mois_de_facturation.month]}{self.mois_de_facturation.year}.pdf"
        pv_to_pdf = os.path.join(enregistrer_sous, nom_du_fichier)
        sheet.range("A1:i53").api.ExportAsFixedFormat(0, pv_to_pdf)
        wk.close()
        return pv_to_pdf

