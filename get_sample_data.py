import  pandas as pd

def read_excel(path):
    all_data = pd.DataFrame()
    try:
        xls = pd.ExcelFile(path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            all_data = pd.concat([all_data, df], ignore_index=True)
            positive_rate = (all_data["Sentiment"] == "Positive").sum() / len(all_data)
            negative_rate = (all_data["Sentiment"] == "Negative").sum() / len(all_data)
            neutral_rate = 1 - positive_rate - negative_rate
            sentiment_rates = {
                "Positive": positive_rate,
                "Negative": negative_rate,
                "Neutral": neutral_rate
            }
        return all_data, sentiment_rates
    except Exception as e:
        print(f"Error processing Excel files: {e}")
    return all_data, {}

def get_sample_data(df, sample_size, sentiment_rates):    
    topics = df["Topic"].unique().tolist()
    try:
        sampled_df = pd.DataFrame()
        for topic in topics:
            sub_df = df[df["Topic"] == topic].reset_index(drop=True)
            if sample_size:
                positive_count = int(sentiment_rates["Positive"] * sample_size)
                negative_count = int(sentiment_rates["Negative"] * sample_size)
                neutral_count = sample_size - positive_count - negative_count
                positive_samples = sub_df[sub_df["Sentiment"] == "Positive"].sample(positive_count, replace=True) if sub_df[sub_df["Sentiment"] == "Positive"].shape[0] > 0 else None
                negative_samples = sub_df[sub_df["Sentiment"] == "Negative"].sample(negative_count, replace=True) if sub_df[sub_df["Sentiment"] == "Negative"].shape[0] > 0 else None
                neutral_samples = sub_df[sub_df["Sentiment"] == "Neutral"].sample(neutral_count, replace=True) if sub_df[sub_df["Sentiment"] == "Neutral"].shape[0] > 0 else None
                if positive_samples is not None:
                    sampled_df = pd.concat([sampled_df, positive_samples], ignore_index=True)
                if negative_samples is not None:
                    sampled_df = pd.concat([sampled_df, negative_samples], ignore_index=True)
                if neutral_samples is not None:
                    sampled_df = pd.concat([sampled_df, neutral_samples], ignore_index=True)
            else:
                sampled_df = pd.concat([sampled_df, sub_df], ignore_index=True)
        return sampled_df
    except Exception as e:
        print(f"Error during sampling: {e}")
    return None


