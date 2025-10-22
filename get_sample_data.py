import pandas as pd
import numpy as np
from calculate_sample_size import get_sample_size

def read_excel(path):
    try:
        xls = pd.ExcelFile(path)
        frames = [pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names]
        all_data = pd.concat(frames, ignore_index=True)

        if all_data.empty:
            return all_data
        all_data["Sentiment"] = (
            all_data["Sentiment"]
            .fillna("")                
            .str.strip()                 
            .str.capitalize()           
        )
        return all_data

    except Exception as e:
        raise ValueError(f"Error processing Excel files: {e}")


import pandas as pd

def get_sample_data(df, params_dict):
    try:
        sampled_list = []
        for topic in df["Topic"].unique():
            sub_df = df[df["Topic"] == topic]
            positive_rate = (sub_df["Sentiment"] == "Positive").sum() / len(sub_df)
            negative_rate = (sub_df["Sentiment"] == "Negative").sum() / len(sub_df)
            neutral_rate = (sub_df["Sentiment"] == "Neutral").sum() / len(sub_df)
            empty_rate = (sub_df["Sentiment"] == "").sum() / len(sub_df)
            p, E, confidence = params_dict.get("response_distribution"), params_dict.get("margin"), params_dict.get("confidence")
            sample_size = get_sample_size(p, E, confidence, N=len(sub_df))
            if sample_size > 0:
                counts = {
                    "Positive": round(positive_rate * sample_size),
                    "Negative": round(negative_rate * sample_size),
                    "Neutral": round(neutral_rate * sample_size),
                    "Empty": round(empty_rate * sample_size)
                }
            
                for sentiment in counts.keys():
                    if counts[sentiment] > 0:            
                        subset = sub_df[sub_df["Sentiment"] == sentiment] if sentiment != "Empty" else sub_df[sub_df["Sentiment"] == ""]
                        if len(subset) > 0:
                            count = min(counts[sentiment], len(subset)) # Đảm bảo không lấy mẫu nhiều hơn dữ liệu gốc
                            sampled = subset.sample(count, replace=False)
                            sampled_list.append(sampled)

        selected_rows = pd.concat(sampled_list) if sampled_list else pd.DataFrame()
        df["Sampled"] = np.where(df.index.isin(selected_rows.index), "x", "")
        return df
    
    except Exception as e:
        print(f"Error during sampling: {e}")
        return pd.DataFrame()



