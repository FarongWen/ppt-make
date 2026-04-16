import pandas as pd
import json
from datetime import datetime

def format_date(date_value):
    """Format date values into a readable string format"""
    if pd.isna(date_value):
        return "未知日期"  # Return placeholder for missing dates
    
    if isinstance(date_value, pd.Timestamp):
        return date_value.strftime('%Y年%m月%d日')
    elif isinstance(date_value, str):
        try:
            # Try to parse string dates (might be in different formats)
            if '-' in date_value:
                # Handle datetime strings with time component
                if ':' in date_value:
                    date_obj = datetime.strptime(date_value.split(' ')[0], '%Y-%m-%d')
                else:
                    date_obj = datetime.strptime(date_value, '%Y-%m-%d')
                return date_obj.strftime('%Y年%m月%d日')
            else:
                return date_value
        except:
            return date_value
    elif isinstance(date_value, int):
        # Handle Excel date integers (days since 1900-01-01)
        try:
            date_obj = datetime(1899, 12, 30) + pd.Timedelta(days=date_value)
            return date_obj.strftime('%Y年%m月%d日')
        except:
            return str(date_value)
    else:
        return str(date_value)

def main():
    try:
        # Read the Excel file
        excel_file = "../副本1-发展对象选拔报名_20260415120238.xlsx"
        df = pd.read_excel(excel_file, sheet_name="Sheet1")
        print(f"Successfully read {len(df)} records from Excel")
        
        # Read the template
        with open("requirement.md", "r", encoding="utf-8") as f:
            template_content = f.read()
        
        # Extract the template part (after "# template")
        template = template_content.split("# template")[1].strip()
        print(f"Template loaded: {template[:50]}...")
        
        # List to store all processed entries
        results = []
        
        # Process each row in the dataframe
        for index, row in df.iterrows():
            try:
                # Create a copy of the template for this member
                member_text = template
                
                # Map of placeholder to column name
                placeholder_map = {
                    "{姓名}": "姓名",
                    "{性别}": "性别",
                    "{民族}": "民族",
                    "{籍贯}": "籍贯",
                    "{班级}": "班级",
                    "{出生日期}": "出生年月",
                    "{入党申请书时间}": "提交入党申请书时间（以实际材料情况为准）",
                    "{确定积极分子日期}": "确定积极分子日期",
                    "{确定为发展对象时间}": "确定为发展对象时间",
                    "{入党日期}": "接收预备党员支部大会时间",
                    "{所属党支部}": "老支部名称",
                    "{学工经历}":"学工经历",
                    "{获奖经历}":"获奖经历"
                }
                
                # Replace each placeholder with the corresponding value
                for placeholder, column in placeholder_map.items():
                    if column in df.columns:
                        value = row[column]
                        # Format dates for better readability
                        if "日期" in column or "时间" in column:
                            formatted_value = format_date(value)
                        else:
                            formatted_value = str(value)
                        
                        member_text = member_text.replace(placeholder, formatted_value)
                
                # Post-processing to fix any remaining timestamp formats
                # Replace patterns like "2025-12-02 00:00:00" with "2025年12月02日"
                import re
                pattern = r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}'
                matches = re.findall(pattern, member_text)
                
                for match in matches:
                    try:
                        date_part = match.split(' ')[0]  # Extract just the date part
                        year, month, day = date_part.split('-')
                        formatted_date = f"{year}年{month}月{day}日"
                        member_text = member_text.replace(match, formatted_date)
                    except:
                        # If parsing fails, leave as is
                        pass
                        
                # Handle NaT values
                member_text = member_text.replace("NaT", "未知日期")
                
                # Fix any remaining timestamp formats that might have been missed
                import re
                timestamp_pattern = r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}'
                
                def timestamp_replacer(match):
                    timestamp_str = match.group(0)
                    try:
                        date_obj = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                        return date_obj.strftime('%Y年%m月%d日')
                    except:
                        return timestamp_str
                
                member_text = re.sub(timestamp_pattern, timestamp_replacer, member_text)
                
                # Add to results
                results.append({
                    "id": index + 1,
                    "name": row["姓名"] if "姓名" in df.columns else f"Party Member {index+1}",
                    "text": member_text
                })
                
                if index < 2:  # Print first 2 entries as samples
                    print(f"\nSample {index+1}:")
                    print(member_text[:100] + "...")
                
            except Exception as e:
                print(f"Error processing row {index+1}: {e}")
        
        # Save results to Excel file
        output_file = "副本1-发展对象选拔报名_20260415120238-整理后.xlsx"
        
        # Create a DataFrame with the text content in the first column only
        df_output = pd.DataFrame({
            "党员信息": [item["text"] for item in results]
        })
        
        # Save to Excel
        df_output.to_excel(output_file, index=False)
        
        print(f"\nProcessed {len(results)} party members")
        print(f"Results saved to {output_file}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
