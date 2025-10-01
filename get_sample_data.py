import  pandas as pd

import pandas as pd

def read_excel(path):
    try:
        xls = pd.ExcelFile(path)
        frames = [pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names]
        all_data = pd.concat(frames, ignore_index=True)

        if all_data.empty:
            return all_data, {}

        positive_rate = (all_data["Sentiment"] == "Positive").sum() / len(all_data)
        negative_rate = (all_data["Sentiment"] == "Negative").sum() / len(all_data)
        neutral_rate = (all_data["Sentiment"] == "Neutral").sum() / len(all_data)
        empty_rate = 1 - (positive_rate + negative_rate + neutral_rate)

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
        for topic in df["Topic"].unique():
            sub_df = df[df["Topic"] == topic]
            counts = {
                "Positive": int(sentiment_rates["Positive"] * sample_size),
                "Negative": int(sentiment_rates["Negative"] * sample_size),
                "Neutral": int(sentiment_rates["Neutral"] * sample_size),
            }
            counts["Empty"] = sample_size - counts["Positive"] - counts["Negative"] - counts["Neutral"]

            for sentiment in sentiments:
                subset = sub_df[sub_df["Sentiment"] == sentiment]
                if not subset.empty and counts[sentiment] > 0:
                    sampled = subset.sample(counts[sentiment], replace=True)
                    sampled_list.append(sampled)
            if counts["Empty"] > 0:
                empty_subset = sub_df[sub_df["Sentiment"].isnull() | (sub_df["Sentiment"] == "")]
                if not empty_subset.empty:
                    sampled_empty = empty_subset.sample(counts["Empty"], replace=True)
                    sampled_list.append(sampled_empty)

        return pd.concat(sampled_list, ignore_index=True) if sampled_list else pd.DataFrame()

    except Exception as e:
        print(f"Error during sampling: {e}")
        return pd.DataFrame()



