import xlwings as xw
import numpy as np


def localise_ref(data, ref):
    """get an elem of a matrix"""
    loc = np.where(data == ref)[0]
    return loc

class DATA:

    def __init__(self, fichier_source, plage_de_localisation, plage_de_données):
        self.fichier_source = fichier_source
        self.plage_de_localisation = plage_de_localisation
        self.plage_de_données=plage_de_données

    def recupere_data(self):
        """get the data in the specified file "fichier_source" """

        wk = xw.Book(self.fichier_source)
        sheet = wk.sheets("TBD")

        # choice of the reference and the minute date.
        mois_de_facturation = sheet.range("b2").value  # return 0
        localisation_mois_de_facturation= mois_de_facturation.month  # return 1
        ref_client = sheet.range("c2").value  # return 2

        # get the location of the reference in the sheet.
        range_localisation = sheet.range(self.plage_de_localisation).value
        range_localisation_array = np.array(range_localisation)
        localisation_ref = localise_ref(range_localisation_array, ref_client)[0]  # return 3

        # define the data range in the sheet.
        range_data = sheet.range(self.plage_de_données).value
        range_data_array = np.array(range_data)

        # Customer name.
        nom_client = range_data_array[localisation_ref, 1]  # return 4
        # Order number.
        num_commande = range_data_array[localisation_ref, 2]  # return 5
        # Project reference.
        ref_projet = range_data_array[localisation_ref, 3]  # return 6
        # Bill number.
        num_facture = range_data_array[localisation_ref, 4]  # return 7
        # Average daily rate (Taux journalier moyen).
        tjm = range_data_array[localisation_ref, 5]  # return 8
        # Invoice description.
        desc_fact = range_data_array[localisation_ref, 7]  # return 9
        # Email address of the recipient.
        destinataire = range_data_array[localisation_ref, 8]  # return 10
        # Supplier reference.
        ref_fournisseur = range_data_array[localisation_ref, 10]  # return 11
        # Batch number.
        num_lot = range_data_array[localisation_ref, 11]  # return 12
        # Number of days ordered (prévues).
        nb_jours_commandés = range_data_array[localisation_ref+5, localisation_mois_de_facturation]  # return 13
        # Number of days billed (réalisés)
        nb_jours_facturés = range_data_array[localisation_ref+7, localisation_mois_de_facturation]  # return 14
        # Discount title.
        intitulé_remise = range_data_array[localisation_ref+12, localisation_mois_de_facturation]  # return 15
        # Discount price.
        prix_unitaire_remise = range_data_array[localisation_ref+13, localisation_mois_de_facturation]  # return 16
        # Discount unit number.
        nb_unité_remise = range_data_array[localisation_ref+14, localisation_mois_de_facturation]  # return 17
        # Discount amount.
        montant_remise = range_data_array[localisation_ref+15, localisation_mois_de_facturation]  # return 18
        # Additional costs title.
        intitulé_frais_annexe = range_data_array[localisation_ref+16, localisation_mois_de_facturation]  # return 19
        # Additional costs price.
        prix_unitaire_frais_annexe = range_data_array[localisation_ref+17, localisation_mois_de_facturation]  # return 20
        # Additionnal cost unit number.
        nb_unité_frais_annexe = range_data_array[localisation_ref+18, localisation_mois_de_facturation]  # return 21
        # Additional costs amount.
        montant_frais_annexe = range_data_array[localisation_ref+19, localisation_mois_de_facturation]  # return 22
        # Email subject.
        sujet_email = sheet.range("j3").value  # return 23
        # Email content.
        contenu_email = sheet.range("i5").value  # return 24


        return (mois_de_facturation,  # return 0
                localisation_mois_de_facturation,  # return 1
                ref_client,  # return 2
                localisation_ref,  # return 3
                nom_client,  # return 4
                num_commande,  # return 5
                ref_projet,  # return 6
                num_facture,  # return 7
                tjm,  # return 8
                desc_fact,  # return 9
                destinataire,  # return 10
                ref_fournisseur,  # return 11
                num_lot,  # return 12
                nb_jours_commandés,  # return 13
                nb_jours_facturés,  # return 14
                intitulé_remise,  # return 15
                prix_unitaire_remise,  # return 16
                nb_unité_remise,  # return 17
                montant_remise,  # return 18
                intitulé_frais_annexe,  # return 19
                prix_unitaire_frais_annexe,  # return 20
                nb_unité_frais_annexe,  # return 21
                montant_frais_annexe,  # return 22
                sujet_email,  # return 23
                contenu_email  # return 24

                )

