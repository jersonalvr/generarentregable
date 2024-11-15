import streamlit as st
import os
import pyperclip
from st_copy_to_clipboard import st_copy_to_clipboard

def crear_donation_footer(base_dir):
    footer = st.container()
    
    with footer:
        st.markdown("---")
        st.header("☕ Buy me a coffee")
        
        # Tabs for different payment methods
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Yape", "Depósito Bancario", "Tarjeta", "Otros Métodos", "Crypto"])
        
        # Yape Tab
        with tab1:
            st.subheader("Donar por Yape")
            col1, col2 = st.columns([1, 1])
            with col1:
                yape_image_path = os.path.join(base_dir, "yape.png")
                if os.path.exists(yape_image_path):
                    st.image(yape_image_path, width=300)
                else:
                    st.error(f"No se encontró la imagen en: {yape_image_path}")
            with col2:
                # Agregamos el número de Yape con botón para copiar
                st_copy_to_clipboard("964536063", "Copiar número de Yape")

        # Bank Deposits Tab
        with tab2:
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.subheader("En Soles 🇵🇪")
                banks_soles = {
                    "BCP": "31004315283063",
                    "BanBif": "008023869670",
                    "Interbank": "8983223094904",
                    "Scotiabank": "7508188435"
                }
                
                for bank, account in banks_soles.items():
                    st.write(f"**{bank}:**")
                    st_copy_to_clipboard(account, f"Copiar cuenta {bank}")
            
            with col2:
                st.subheader("En Dólares 💵")
                banks_usd = {
                    "BCP": "31004319160179",
                    "Interbank": "8983224537574",
                    "Scotiabank": "7508188145"
                }
                
                for bank, account in banks_usd.items():
                    st.write(f"**{bank}:**")
                    st_copy_to_clipboard(account, f"Copiar cuenta {bank}")
        
        # Card Payments Tab
        with tab3:
            st.subheader("Donar con tarjeta 💳")
            
            amounts = {
                "10": "https://pago-seguro.vendemas.com.pe/MTYzNDc3OTY3NjYxMWM4MjU2MTIuNzMxNzMxMjgxNTUz",
                "15": "https://pago-seguro.vendemas.com.pe/ZjE5MTc2MWIzMjM0MDQ3NDQ0NC4yOWQxNzMxMjgxNjIz",
                "20": "https://pago-seguro.vendemas.com.pe/NmEzOTkwNzI1OTY1Zi42MTM0MDg2NDYxNzMxMjgxNjUy",
                "25": "https://pago-seguro.vendemas.com.pe/MTUzODVjYTM4NDIzZDgxNjMwLjI0NzcxNzMxMjgxNjc3",
                "30": "https://pago-seguro.vendemas.com.pe/ODM0LjIyMTYyMjA2NjZmYjdhNmMzM2QxNzMxMjgxNzA3",
                "35": "https://pago-seguro.vendemas.com.pe/ODM3MzMzMDFkNDcuOTE2YzUzMjQxNTExNzMxMjgxNzI1",
                "40": "https://pago-seguro.vendemas.com.pe/YzgwZjM2NzM1LjZiMWQ0NzEzNTMxNzYxNzMxMjgxNzQ0",
                "45": "https://pago-seguro.vendemas.com.pe/MTgwMTYuNzQ1MzM0OTk0MzI4MTQ2MzYxNzMxMjgxNzYy"
            }
            
            # Creamos dos filas de 4 columnas cada una para mejor visualización
            for row in range(2):
                cols = st.columns(4)
                start_idx = row * 4
                end_idx = start_idx + 4
                
                # Tomamos solo los montos correspondientes a esta fila
                row_amounts = dict(list(amounts.items())[start_idx:end_idx])
                
                for col_idx, (amount, link) in enumerate(row_amounts.items()):
                    with cols[col_idx]:
                        st.link_button(f"S/ {amount}", link)
            
            st.link_button("Más de S/ 50", "https://linkdecobro.ligo.live/v3/44df73097f594239b21b78b6905bed98")
        
        # Other Payment Methods Tab
        with tab4:
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.subheader("Mercado Pago")
                st.link_button("Donar con Mercado Pago", "https://link.mercadopago.com.pe/jersonapp")
            
            with col2:
                st.subheader("PayPal")
                st.link_button("Donar con PayPal", "https://www.paypal.com/paypalme/dschimbote")
        
        # Crypto Tab
        with tab5:
            st.subheader("Binance")
            col1, col2 = st.columns([1, 1])
            
            with col1:
                binance_image_path = os.path.join(base_dir, "binance.png")
                if os.path.exists(binance_image_path):
                    st.image(binance_image_path, width=300)
                else:
                    st.error(f"No se encontró la imagen en: {binance_image_path}")
            
            with col2:
                st.link_button("Donar con Binance", "https://app.binance.com/qr/dplkbb7f88c5329c4692adf278670d1b37ab")
