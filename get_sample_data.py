import pandas as pd
import numpy as np

def read_excel(path):
    try:
        xls = pd.ExcelFile(path)
        frames = [pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names]
        all_data = pd.concat(frames, ignore_index=True)

        if all_data.empty:
            return all_data, {}
        all_data["Sentiment"] = (
            all_data["Sentiment"]
            .fillna("")                
            .str.strip()                 
            .str.capitalize()           
        )
        positive_rate = (all_data["Sentiment"] == "Positive").sum() / len(all_data)
        negative_rate = (all_data["Sentiment"] == "Negative").sum() / len(all_data)
        neutral_rate = (all_data["Sentiment"] == "Neutral").sum() / len(all_data)
        empty_rate = (all_data["Sentiment"] == "").sum() / len(all_data)

        sentiment_rates = {
            "Positive": positive_rate,
            "Negative": negative_rate,
            "Neutral": neutral_rate,
            "Empty": empty_rate
        }
        return all_data, sentiment_rates

    except Exception as e:
        raise ValueError(f"Error processing Excel files: {e}")


import pandas as pd

def get_sample_data(df, sample_size, sentiment_rates):
    try:
        if not sample_size:
            return df.copy()
        sampled_list = []
        sentiments = ["Positive", "Negative", "Neutral"]
        
        sub_df = df[["Sentiment"]]
        counts = {
            "Positive": int(sentiment_rates["Positive"] * sample_size),
            "Negative": int(sentiment_rates["Negative"] * sample_size),
            "Neutral": int(sentiment_rates["Neutral"] * sample_size),
            "Empty": int(sentiment_rates["Empty"] * sample_size)
        }
        
        for sentiment in sentiments:
            if counts[sentiment] > 0:            
                subset = sub_df[sub_df["Sentiment"] == sentiment]
                if not subset.empty:
                    sampled = subset.sample(counts[sentiment], replace=False)
                    sampled_list.append(sampled)
                   
        if counts["Empty"] > 0:
            empty_subset = sub_df[sub_df["Sentiment"] == ""]
            if len(empty_subset) > 0:
                sampled_empty = empty_subset.sample(counts["Empty"], replace=False)
                sampled_list.append(sampled_empty)
    

        selected_rows = pd.concat(sampled_list) if sampled_list else pd.DataFrame()
        df["Sampled"] = np.where(df.index.isin(selected_rows.index), "x", "")
        return df
    
    except Exception as e:
        print(f"Error during sampling: {e}")
        return pd.DataFrame()



