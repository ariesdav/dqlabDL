import pandas as pd
from mlxtend.frequent_patterns import apriori,  association_rules

def run_analysis(input_xlsx_path: str, output_xlsx_path: str) -> None: 
    # baca file dan sheet "transaksi"
    df = pd.read_excel(input_xlsx_path, sheet_name="Transaksi", dtype=str)
    # ubah product ke group_product
    group_product=(df.groupby(['Kode Transaksi','Nama Produk'])['Jumlah'].count().unstack().fillna(0))
    # encode
    group_product=group_product.apply(lambda col: col.map(lambda x: 1 if x > 0 else 0))
    # apriori
    frequent_itemsets=apriori(group_product,min_support=0.05,use_colnames=True)
    # association rules
    rules=association_rules(frequent_itemsets,metric='confidence',min_threshold=0.4)
    # nama produk sesuai alfabet
    rules["itemset_sorted"]=rules.apply(lambda x: frozenset(sorted(list(x["antecedents"]|x["consequents"]))),axis=1)
    # hilangkan duplikat
    rules=rules.drop_duplicates(subset=["itemset_sorted"])
    # urutkan lift desc lalu confidence desc
    rules=rules.sort_values(by=["lift","confidence"], ascending=[False,False]).reset_index(drop=True)
    # menyesuaikan format output 
    output_df=pd.DataFrame({
        'Packaging Set ID': range(1, len(rules)+1),
        'Products': rules.apply(lambda x: ';'.join(sorted(list(x["antecedents"]|x["consequents"]))),axis=1),
        'Maximum Lift': rules['lift'].round(2), 
        'Maximum Confidence': rules['confidence'].round(2)
    })
    # simpan ke excel
    with pd.ExcelWriter(output_xlsx_path, engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name="Packaging", index=False)
run_analysis("transaksi_dqmart.xlsx", "product_packaging.xlsx")