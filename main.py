from datetime import datetime
import os
import streamlit as st
from traitement_pv import*
from get_data import*

MAINTPATH = r"C:\Users\jrala\OneDrive\BienCommun"
TEMPLATE = os.path.join(MAINTPATH, "pv.xlsm")
date_string = '2021-12-31'
MOIS_DE_FACTURATION = datetime.strptime(date_string, '%Y-%m-%d')

def execute():

    data = DATA(fichier_source=fichier_source,
                plage_de_localisation=plage_de_localisation,
                plage_de_données=plage_de_données)
    list_data_recuperé = data.recupere_data()

    pv = PV(ref_client=list_data_recuperé[2],
            nom_client=list_data_recuperé[4],
            num_commande=list_data_recuperé[5],
            ref_projet=list_data_recuperé[6],
            num_facture=list_data_recuperé[7],
            tjm=list_data_recuperé[8],
            desc_fact=list_data_recuperé[9],
            destinataire=list_data_recuperé[10],
            ref_fournisseur=ref_fournisseur,
            num_lot=list_data_recuperé[12],
            nb_jours_commandés=list_data_recuperé[13],
            nb_jours_facturés=list_data_recuperé[14],
            mois_de_facturation=list_data_recuperé[0],
            intitulé_remise=list_data_recuperé[15],
            prix_unitaire_remise=list_data_recuperé[16],
            nb_unité_remise=list_data_recuperé[17],
            montant_remise=list_data_recuperé[18],
            intitulé_frais_annexe=list_data_recuperé[19],
            prix_unitaire_frais_annexe=list_data_recuperé[20],
            nb_unité_frais_annexe=list_data_recuperé[21],
            montant_frais_annexe=list_data_recuperé[22],
            sujet_email=list_data_recuperé[23],
            contenu_email=list_data_recuperé[24]
            )

    pv.rempli_pv(TEMPLATE)
    to_pdf = pv.convert_to_pdf(enregistrer_sous, TEMPLATE)
    st.write(to_pdf)
    pv.send_mail(to_pdf)


if __name__ == "__main__":

    st.title("Edition de PV")
    fichier_source = st.text_input("NOM DU FICHIER SOURCE", value="Dashboard.xlsm")
    plage_de_localisation= st.text_input("PLAGE DE LOCALISATION", value="b11:b2000")
    plage_de_données = st.text_input("PLAGE DE DONNEES", value="b11:n2000")
    ref_fournisseur = st.text_input("REFERENCE FOURNISSEUR", value="0000018824")
    enregistrer_sous = st.text_input("ENREGISTRER PV A SIGNER DANS LE DOSSIER")

    if st.button('executer'):
        execute()
