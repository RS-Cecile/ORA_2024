import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
import openpyxl
import xlsxwriter
from io import BytesIO

st.set_page_config(
    page_title="ORA 2024 - Par tranche d'effectif salarié",
    page_icon=	":chart_with_upwards_trend:",
    layout="wide")

st.markdown("# Résultats par tranche d'effectif salarié")
st.sidebar.markdown("# Par tranches d'effectif salarié")

fichier = "ORA_donnee.xlsx"
sheet = "Salariés"

############################################################

# Votre association prend-elle en compte les enjeux liés à la transition écologique pour mener à bien ses activités et organiser son action ?
table = pd.read_excel( fichier, sheet_name = sheet ,skiprows=11,nrows= 8, index_col =0, dtype = "object")
table = table.applymap(lambda x: f'{x * 100:.0f}%')

# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#Les économies d'énergie (électricité, gaz,...) et de la ressource en eau
table1 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=89,nrows= 5, index_col =0, dtype = "object")
table1 = table1.applymap(lambda x: f'{x * 100:.0f}%')


# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#La limitation des déplacements, les transports collectifs et les mobilités douces (vélo…)
table2 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=97,nrows= 5, index_col =0, dtype = "object")
table2 = table2.applymap(lambda x: f'{x * 100:.0f}%')

# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#La gestion des déchets (tri sélectif, moins d'emballage, biodéchets...)
table3 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=105,nrows= 5, index_col =0, dtype = "object")
table3 = table3.applymap(lambda x: f'{x * 100:.0f}%')

# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#Des achats responsables (en local, circuit-court...)
table4 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=113,nrows= 5, index_col =0, dtype = "object")
table4 = table4.applymap(lambda x: f'{x * 100:.0f}%')

# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#Le recours à des fournitures plus écologiques (papier recyclé, cartouches d'encre rechargeables...)
table5 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=121,nrows= 5, index_col =0, dtype = "object")
table5 = table5.applymap(lambda x: f'{x * 100:.0f}%')

# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#Le réemploi, le recours aux recycleries et aux entreprises d'insertion à vocation environnementale
table6 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=129,nrows= 5, index_col =0, dtype = "object")
table6 = table6.applymap(lambda x: f'{x * 100:.0f}%')

# Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
#La sobriété numérique (utilisation durable et raisonnable du numérique)
table7 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=137,nrows= 5, index_col =0, dtype = "object")
table7 = table7.applymap(lambda x: f'{x * 100:.0f}%')

# Qu'est-ce qui pourrait aider votre association à [mieux] prendre en compte les enjeux liés à la transition écologique dans ses activités et son fonctionnement ? Plusieurs réponses possibles
table8 = pd.read_excel( fichier, sheet_name = sheet ,skiprows=148,nrows= 9, index_col =0, dtype = "object")
table8 = table8.applymap(lambda x: f'{x * 100:.0f}%')

styled_table = table.style.set_properties(**{'text-align': 'center'})
styled_table1 = table1.style.set_properties(**{'text-align': 'center'})
styled_table2 = table2.style.set_properties(**{'text-align': 'center'})
styled_table3 = table3.style.set_properties(**{'text-align': 'center'})
styled_table4 = table4.style.set_properties(**{'text-align': 'center'})
styled_table5 = table5.style.set_properties(**{'text-align': 'center'})
styled_table6 = table6.style.set_properties(**{'text-align': 'center'})
styled_table7 = table7.style.set_properties(**{'text-align': 'center'})
styled_table8 = table8.style.set_properties(**{'text-align': 'center'})

tab1, tab2, tab3 = st.tabs(["Ensemble des tranches d'effectif salarié ","Par tranche d'effectif salarié",  "Téléchargement des données"])

with tab1:
    st.header("Votre association prend-elle en compte les enjeux liés à la transition écologique pour mener à bien ses activités et organiser son action ?")
    st.table(styled_table)
    st.header("Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?")
    st.subheader("Les économies d'énergie (électricité, gaz,...) et de la ressource en eau")
    st.table(styled_table1)
    st.subheader("La limitation des déplacements, les transports collectifs et les mobilités douces (vélo…)")
    st.table(styled_table2)
    st.subheader("La gestion des déchets (tri sélectif, moins d'emballage, biodéchets...)")
    st.table(styled_table3)
    st.subheader("Des achats responsables (en local, circuit-court...)")
    st.table(styled_table4)
    st.subheader("Le recours à des fournitures plus écologiques (papier recyclé, cartouches d'encre rechargeables...)")
    st.table(styled_table5)
    st.subheader("Le réemploi, le recours aux recycleries et aux entreprises d'insertion à vocation environnementale")
    st.table(styled_table6)
    st.subheader("La sobriété numérique (utilisation durable et raisonnable du numérique)")
    st.table(styled_table7)
    st.header("Qu'est-ce qui pourrait aider votre association à [mieux] prendre en compte les enjeux liés à la transition écologique dans ses activités et son fonctionnement ? *Plusieurs réponses possibles*")
    st.table(styled_table8)
    
with tab2:
    option = st.selectbox(
        "**Sélectionnez la tranche d'effectif salarié souhaitée, les résultats correspondants apparaîtront dans la colonne bleue du tableau et sous forme de graphiques :**",
        ("1 ou 2 salariés",
         "3 à 5 salariés",
         "6 à 9 salariés",
         "10 à 19 salariés",
         "20 à 49 salariés"))

    st.markdown(f'<h1 style = "font-size:30px;"> Vous avez sélectionné : </h1>'f'<h1 style = "font-size:40px;color:royalblue;text-transform:uppercase;"> {option} </h1>', unsafe_allow_html=True)

    # Votre association prend-elle en compte les enjeux liés à la transition écologique pour mener à bien ses activités et organiser son action ?
    table_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=11,nrows= 8, index_col =0, dtype = "object")
    table_ = table_.applymap(lambda x: f'{x * 100:.0f}%')
    slice = [option]
    forme = table.style.set_properties(**{'background-color' : 'cornflowerblue', 'text-align': 'center', 'color' : "white"}, subset = slice)
    st.header(
        "Votre association prend-elle en compte les enjeux liés à la transition écologique pour mener à bien ses activités et organiser son action ?")
    st.table(forme)


    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #Les économies d'énergie (électricité, gaz,...) et de la ressource en eau
    table1_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=89,nrows= 3, index_col =0, dtype = "object")
    fig_col = table1_[option] * 100
    names = fig_col.index
    st.header(
        "Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?")
    st.subheader("Les économies d'énergie (électricité, gaz,...) et de la ressource en eau")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #La limitation des déplacements, les transports collectifs et les mobilités douces (vélo…)
    table2_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=97,nrows= 3, index_col =0, dtype = "object")
    fig_col = table2_[option] * 100
    names = fig_col.index
    st.subheader("La limitation des déplacements, les transports collectifs et les mobilités douces (vélo…)")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #La gestion des déchets (tri sélectif, moins d'emballage, biodéchets...)
    table3_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=105,nrows= 3, index_col =0, dtype = "object")
    fig_col = table3_[option] * 100
    names = fig_col.index
    st.subheader("La gestion des déchets (tri sélectif, moins d'emballage, biodéchets...)")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #Des achats responsables (en local, circuit-court...)
    table4_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=113,nrows= 3, index_col =0, dtype = "object")
    fig_col = table4_[option] * 100
    names = fig_col.index
    st.subheader("Des achats responsables (en local, circuit-court...)")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #Le recours à des fournitures plus écologiques (papier recyclé, cartouches d'encre rechargeables...)
    table5_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=121,nrows= 3, index_col =0, dtype = "object")
    fig_col = table5_[option] * 100
    names = fig_col.index
    st.subheader("Le recours à des fournitures plus écologiques (papier recyclé, cartouches d'encre rechargeables...)")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #Le réemploi, le recours aux recycleries et aux entreprises d'insertion à vocation environnementale
    table6_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=129,nrows= 3, index_col =0, dtype = "object")
    fig_col = table6_[option] * 100
    names = fig_col.index
    st.subheader("Le réemploi, le recours aux recycleries et aux entreprises d'insertion à vocation environnementale")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    # Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?
    #La sobriété numérique (utilisation durable et raisonnable du numérique)
    table7_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=137,nrows= 3, index_col =0, dtype = "object")
    fig_col = table7_[option] * 100
    names = fig_col.index
    st.subheader("La sobriété numérique (utilisation durable et raisonnable du numérique)")
    fig = px.pie(fig_col, values=option, names=names, color=names, color_discrete_map={"Beaucoup d'attention": "darkblue",
                                                                                       "Une attention modérée": "royalblue",
                                                                                       "Peu ou pas": "lightcyan"})
    st.plotly_chart(fig, use_container_width=True)

    liste = []

    for tab in (table1_, table2_, table3_, table4_, table5_, table6_, table7_):
        derniere_ligne = tab[option].iloc[-1]*100
        liste.append(derniere_ligne)

    index = ["Les économies d'énergie (électricité, gaz,...) et de la ressource en eau",
                "La limitation des déplacements, les transports collectifs et les mobilités douces (vélo…)",
                "La gestion des déchets (tri sélectif, moins d'emballage, biodéchets...)",
                "Des achats responsables (en local, circuit-court...)",
                "Le recours à des fournitures plus écologiques (papier recyclé, cartouches d'encre rechargeables...)",
                "Le réemploi, le recours aux recycleries et aux entreprises d'insertion à vocation environnementale",
                "La sobriété numérique (utilisation durable et raisonnable du numérique)"]
    df1 = pd.DataFrame({"Beaucoup d'attention" : liste}, index = index)
    df1.sort_values(by= "Beaucoup d'attention", ascending= True, inplace = True)
    fig, ax = plt.subplots()
    df1.plot(kind='barh', ax=ax)
    st.subheader("**Synthèse des réponses « Beaucoup d'attention »**")
    ax.set_xlabel("Valeurs (en %)")
    st.pyplot(fig)

    # Qu'est-ce qui pourrait aider votre association à [mieux] prendre en compte les enjeux liés à la transition écologique dans ses activités et son fonctionnement ? Plusieurs réponses possibles
    table8_ = pd.read_excel( fichier, sheet_name = sheet ,skiprows=148,nrows= 9, index_col =0, dtype = "object")
    fig_col = table8_[option]*100
    fig_col.sort_values(ascending= True, inplace = True)
    fig, ax = plt.subplots()
    fig_col.plot(kind='barh', ax = ax)
    for index, value in enumerate(fig_col):
        ax.text(value, index, f'{value:.0f}%', va='center')
    ax.set_xlabel("Valeurs (en %)")
    st.header("Qu'est-ce qui pourrait aider votre association à [mieux] prendre en compte les enjeux liés à la transition écologique dans ses activités et son fonctionnement ? *Plusieurs réponses possibles*")
    st.pyplot(fig)

with tab3 :
    #liste des questions pour l'onglet 2:
    questions = ["Les économies d'énergie (électricité, gaz,...) et de la ressource en eau",
                 "La limitation des déplacements, les transports collectifs et les mobilités douces (vélo…)",
                 "La gestion des déchets (tri sélectif, moins d'emballage, biodéchets...)",
                 "Des achats responsables (en local, circuit-court...)",
                 "Le recours à des fournitures plus écologiques (papier recyclé, cartouches d'encre rechargeables...)",
                 "Le réemploi, le recours aux recycleries et aux entreprises d'insertion à vocation environnementale",
                 "La sobriété numérique (utilisation durable et raisonnable du numérique)"]

    # Liste des tables
    tables = [table1, table2, table3, table4, table5, table6, table7]


    # Fonction pour créer un fichier Excel avec les questions et les tables
    def to_excel(table, tables, questions, table8):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Formats
            bold_format = workbook.add_format({'bold': True, 'text_wrap': True})
            border_format = workbook.add_format({'border': 1, 'text_wrap': True})
            wrap_format = workbook.add_format({'text_wrap': True})
            title_format = workbook.add_format({'bold': True, 'font_size': 14})

            #onglet 1
            worksheet1 = workbook.add_worksheet("Prise en compte des enjeux")
            worksheet1.write(0, 0, "Votre association prend-elle en compte les enjeux liés à la transition écologique pour mener à bien ses activités et organiser son action ?", title_format)
            for col_idx, col_name in enumerate(table.columns):
                worksheet1.write(2, col_idx+1, col_name, bold_format)
                worksheet1.set_column(col_idx+1, col_idx, 20)
            for row_idx, data_row in enumerate(table.itertuples(index=True)):
                for col_idx, value in enumerate(data_row):
                    worksheet1.write(row_idx + 3, col_idx, value, border_format)

            # Onglet 2 : Tables 1 à 7
            worksheet2 = workbook.add_worksheet('Pratiques')
            worksheet2.write(0, 0,
                             "Quelle attention porte votre association aux pratiques suivantes dans la conduite de ses activités et dans son organisation ?",
                             title_format)
            row = 3
            for title, table in zip(questions, tables):
                worksheet2.write(row, 0, title, title_format)
                row += 1
                for col_idx, col_name in enumerate(table.columns):
                    worksheet2.write(row + 1, col_idx + 1, col_name, bold_format)
                    worksheet2.set_column(col_idx+1, col_idx, 20)
                for data_row in table.itertuples(index=True):
                    row += 1
                    for col_idx, value in enumerate(data_row):
                        worksheet2.write(row + 1, col_idx, value, border_format)
                row += 3

            # Onglet 3 : Un seul tableau avec son titre
            worksheet3 = workbook.add_worksheet('Aides pour la prise en compte')
            worksheet3.write(0, 0, "Qu'est-ce qui pourrait aider votre association à [mieux] prendre en compte les enjeux liés à la transition écologique dans ses activités et son fonctionnement ? *Plusieurs réponses possibles*", title_format)
            for col_idx, col_name in enumerate(table8.columns):
                worksheet3.write(2, col_idx + 1, col_name, bold_format)
                worksheet3.set_column(col_idx+1, col_idx, 20)
            for row_idx, data_row in enumerate(table8.itertuples(index=True)):
                for col_idx, value in enumerate(data_row):
                    worksheet3.write(row_idx + 3, col_idx, value, border_format)

        return output.getvalue()


    # Bouton de téléchargement
    tables = [table1, table2, table3, table4, table5, table6, table7]
    excel_data = to_excel(table, tables, questions, table8)
    "Pour télécharger les données, cliquez sur le bouton."
    st.download_button(label="Télécharger les données", data=excel_data, file_name="ORA2024-Transition_ecologique-Tranches_effectifs.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       help="Cliquez ici pour télécharger les données au format XLSX")
