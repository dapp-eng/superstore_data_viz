import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Superstore Data Visualization", layout="wide")
st.title("Superstore Data Visualization")
st.markdown("##### by: Daffa Ahmad Pangreksa - 159 - INT24\n##### Data Science UNESA")

try:
    orders = pd.read_excel('superstore_order.xlsx')
    products = pd.read_excel('superstore_product.xlsx')
    customers = pd.read_excel('superstore_customer.xlsx')
    stock = pd.read_excel('product_stock.xlsx')

    def clean_columns(df):
        df.columns = [col.lower().strip().replace(' ', '_').replace('-', '_') for col in df.columns]
        return df

    orders = clean_columns(orders)
    products = clean_columns(products)
    customers = clean_columns(customers)
    stock = clean_columns(stock)

    stock = stock.rename(columns={'stock': 'quantity'})

    if 'order_date' in orders.columns:
        orders['order_date'] = pd.to_datetime(orders['order_date'], errors='coerce')
        orders = orders.dropna(subset=['order_date'])

    numeric_cols = ['sales', 'profit', 'quantity', 'discount']
    for col in numeric_cols:
        if col in orders.columns:
            orders[col] = pd.to_numeric(orders[col], errors='coerce').fillna(0)

    df = orders.merge(customers, on='customer_id', how='left')
    df = df.merge(products, on='product_id', how='left')

    cols_to_drop = []
    cols_to_rename = {}
    
    for col in df.columns:
        if col.endswith('_y'):
            cols_to_drop.append(col)
        elif col.endswith('_x'):
            cols_to_rename[col] = col[:-2]

    df = df.drop(columns=cols_to_drop)
    df = df.rename(columns=cols_to_rename)

    text_cols = ['customer_name', 'segment', 'city', 'state', 'region', 'category', 'sub_category', 'product_name']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna('Unknown')

    df['year_month'] = df['order_date'].dt.to_period('M').astype(str)

    min_date = df['order_date'].min().date()
    max_date = df['order_date'].max().date()

    with st.sidebar:
        st.header("Filters")
        start_date, end_date = st.date_input(
            "Select Date Range",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        
        region_options = ['All'] + sorted(list(df['region'].unique()))
        selected_region = st.selectbox("Select Region", region_options)
        
        category_options = ['All'] + sorted(list(df['category'].unique()))
        selected_category = st.selectbox("Select Category", category_options)

    df_filtered = df[(df['order_date'].dt.date >= start_date) & (df['order_date'].dt.date <= end_date)]
    
    if selected_region != 'All':
        df_filtered = df_filtered[df_filtered['region'] == selected_region]
    
    if selected_category != 'All':
        df_filtered = df_filtered[df_filtered['category'] == selected_category]

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    total_sales = df_filtered['sales'].sum()
    total_profit = df_filtered['profit'].sum()
    total_orders = df_filtered['order_id'].nunique()
    profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0

    kpi1.metric("Total Revenue", f"${total_sales:,.0f}")
    kpi2.metric("Total Profit", f"${total_profit:,.0f}")
    kpi3.metric("Profit Margin", f"{profit_margin:.1f}%")
    kpi4.metric("Total Orders", f"{total_orders:,}")

    st.markdown("---")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("Monthly Sales & Profit Trend")
        monthly_data = df_filtered.groupby('year_month')[['sales', 'profit']].sum().reset_index()
        fig_trend = px.line(monthly_data, x='year_month', y=['sales', 'profit'], 
                            labels={'value': 'Amount', 'year_month': 'Month', 'variable': 'Metric'},
                            markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)

    with col2:
        st.subheader("Profit by Segment")
        segment_data = df_filtered.groupby('segment')['profit'].sum().reset_index()
        fig_donut = px.pie(segment_data, values='profit', names='segment', hole=0.4)
        st.plotly_chart(fig_donut, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Top 10 States by Sales")
        state_data = df_filtered.groupby('state')['sales'].sum().reset_index().sort_values('sales', ascending=True).tail(10)
        fig_bar_state = px.bar(state_data, x='sales', y='state', orientation='h', text_auto='.2s',
                               title="Highest Revenue Generating States")
        st.plotly_chart(fig_bar_state, use_container_width=True)

    with col4:
        st.subheader("Category Performance (Sales vs Profit)")
        cat_data = df_filtered.groupby('sub_category')[['sales', 'profit']].sum().reset_index().sort_values('sales', ascending=False)
        fig_bar_cat = px.bar(cat_data, x='sub_category', y=['sales', 'profit'], barmode='group',
                             title="Revenue vs Profit per Sub-Category")
        st.plotly_chart(fig_bar_cat, use_container_width=True)

    st.subheader("Product Correlation Analysis")
    fig_scatter = px.scatter(df_filtered, x='sales', y='profit', color='category', size='quantity',
                             hover_data=['product_name'], opacity=0.6,
                             title="Sales vs Profit Distribution (Bubble Size = Quantity)")
    st.plotly_chart(fig_scatter, use_container_width=True)

    st.subheader("Low Stock Alert (Inventory)")
    
    stock_merged = products.merge(stock, on='product_id', how='inner')
    
    cols_rename_stock = {}
    for col in stock_merged.columns:
        if col.endswith('_x'):
            cols_rename_stock[col] = col[:-2]
    stock_merged = stock_merged.rename(columns=cols_rename_stock)
    
    if 'product_name' in stock_merged.columns and 'quantity' in stock_merged.columns:
        low_stock = stock_merged[stock_merged['quantity'] < 20].sort_values('quantity').head(10)
        
        if not low_stock.empty:
            fig_stock = px.bar(low_stock, x='quantity', y='product_name', orientation='h',
                               color='quantity', color_continuous_scale='Reds',
                               title="Top Products with Critical Low Stock (<20)")
            st.plotly_chart(fig_stock, use_container_width=True)
        else:
            st.info("Good news! No products are currently running low on stock (< 20 units).")
    else:
        st.warning("Could not find 'quantity' or 'product_name' columns in the stock data.")

except Exception as e:

    st.error(f"Error: {e}")

