import os
import re
from bs4 import BeautifulSoup
from datetime import datetime

def generate_comparison_html(esg_reports):
    """
    Generates an HTML comparison table for up to 5 ESG reports
    :param esg_reports: List of dictionaries containing ESG data from reports
    :return: Filename of the generated HTML file
    """
    if not esg_reports or len(esg_reports) > 5:
        print("❌ Error: Please provide between 1 and 5 reports for comparison")
        return None
    
    # Create safe filename
    output_file = "ESG_Comparison.html"
    current_date = datetime.now().strftime("%B %d, %Y")
    
    # Extract company names and scores
    company_names = [report['company_name'] for report in esg_reports]
    sentiment_scores = [report['sentiment_score'] for report in esg_reports]
    
    # Calculate dynamic column widths
    num_companies = len(company_names)
    first_col_width = "12%"  # Reduced from 30% to 12% (60% reduction)
    other_col_width = f"{(88/num_companies):.2f}%"  # Distribute remaining 88% space
    
    # Generate comparison tables for E, S, G
    def generate_comparison_section(title, icon, category):
        section_html = f"""
            <h2><span class="category-icon">{icon}</span>{title} Comparison</h2>
            <table>
                <thead>
                    <tr>
                        <th width="{first_col_width}">Insight</th>
        """
        
        # Add company headers with dynamic width
        for name in company_names:
            section_html += f'<th width="{other_col_width}">{name}</th>'
        section_html += "</tr></thead><tbody>"
        
        # Find maximum number of insights across all reports for this category
        max_insights = max(len(report[category]) for report in esg_reports)
        
        # Add insights (up to 10 per category)
        for i in range(min(10, max_insights)):
            section_html += f'<tr><td>Insight {i+1}</td>'
            for report in esg_reports:
                # Get the insight if it exists, otherwise use "N/A"
                insight = report[category][i] if i < len(report[category]) else "N/A"
                section_html += f'<td>{insight}</td>'
            section_html += '</tr>'
        
        section_html += "</tbody></table>"
        return section_html
    
    # Build the complete HTML content
    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>ESG Comparison Report</title>
        <style>
            body {{ 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #333;
                background-color: #f9f9f9;
                padding: 0;
                margin: 0;
            }}
            .container {{
                max-width: 1000px;
                margin: 20px auto;
                background: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 0 20px rgba(0,0,0,0.1);
            }}
            header {{
                border-bottom: 2px solid #2196F3;
                padding-bottom: 20px;
                margin-bottom: 30px;
            }}
            h1, h2, h3 {{
                color: #2c3e50;
            }}
            h1 {{
                margin-top: 0;
                font-size: 2.2em;
            }}
            h2 {{
                border-bottom: 1px solid #eee;
                padding-bottom: 8px;
                margin-top: 30px;
                font-size: 1.5em;
                color: #2196F3;
            }}
            h3.subtitle {{
                color: #7f8c8d;
                font-weight: normal;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
                font-size: 0.95em;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 12px 15px;
                text-align: left;
            }}
            th {{
                background-color: #2196F3;
                color: white;
                font-weight: bold;
            }}
            tr:nth-child(even) {{
                background-color: #f2f2f2;
            }}
            tr:hover {{
                background-color: #e3f2fd;
            }}
            .sentiment {{
                font-size: 1.2em;
                padding: 10px 15px;
                background-color: #e8f5e9;
                border-radius: 4px;
                display: inline-block;
                margin: 10px 0;
            }}
            .score-table {{
                margin-bottom: 40px;
            }}
            footer {{
                margin-top: 40px;
                text-align: center;
                font-size: 0.9em;
                color: #7f8c8d;
                border-top: 1px solid #eee;
                padding-top: 20px;
            }}
            .category-icon {{
                font-size: 1.2em;
                margin-right: 8px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1>ESG Comparison Report</h1>
                <h3 class="subtitle">Generated on: {current_date}</h3>
            </header>
            
            <h2>📊 Overall ESG Comparison</h2>
            <table class="score-table">
                <thead>
                    <tr>
                        <th width="30%">Company</th>
                        <th width="70%">ESG Sentiment Score</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    # Add score comparison
    for name, score in zip(company_names, sentiment_scores):
        html_content += f"""
                    <tr>
                        <td>{name}</td>
                        <td>{score}/10</td>
                    </tr>
        """
    
    html_content += """
                </tbody>
            </table>
    """
    
    # Add comparison sections
    html_content += generate_comparison_section("Environmental", "🌍", "environment")
    html_content += generate_comparison_section("Social", "🏢", "social")
    html_content += generate_comparison_section("Governance", "🏛", "governance")
    
    # Footer
    html_content += f"""
            <footer>
                ESG Comparison Report generated on {current_date}<br>
                </footer>
        </div>
    </body>
    </html>
    """
    
    # Save file
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(html_content)
    
    print(f"✅ ESG Comparison Report generated: {output_file}")
    return output_file

def extract_data_from_html(html_file):
    """
    Extracts ESG data from a single HTML report file using BeautifulSoup
    :param html_file: Path to HTML file
    :return: Dictionary with extracted data
    """
    with open(html_file, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
    
    # Extract company name from header
    header = soup.find('header')
    company_name = "Unknown Company"
    if header:
        h1 = header.find('h1')
        if h1:
            company_name = h1.text.replace('ESG Report Analysis', '').strip()
    
    # Extract ticker from subtitle if available
    subtitle = soup.find('h3', class_='subtitle')
    if subtitle:
        subtitle_text = subtitle.get_text()
        ticker_match = re.search(r'Ticker: (\w+)', subtitle_text)
        if ticker_match:
            company_name = f"{company_name} ({ticker_match.group(1)})"
    
    # Extract sentiment score
    sentiment_score = "N/A"
    sentiment_div = soup.find('div', class_='sentiment')
    if sentiment_div:
        score_match = re.search(r'(\d+(?:\.\d+)?)/10', sentiment_div.get_text())
        if score_match:
            sentiment_score = score_match.group(1)
    
    # Function to extract insights from a section
    def extract_insights(section_icon):
        insights = []
        # Find the h2 with the matching icon
        section_header = soup.find(lambda tag: tag.name == 'h2' and section_icon in tag.get_text())
        if section_header:
            # Find the next table
            table = section_header.find_next('table')
            if table:
                # Extract all insight rows (skip header row)
                rows = table.find_all('tr')[1:]  # Skip header row
                for row in rows:
                    cells = row.find_all('td')
                    if len(cells) >= 2:  # Should have at least number and insight
                        insight = cells[1].get_text().strip()
                        insights.append(insight)
        return insights[:10]  # Return max 10 insights
    
    # Extract all insights
    env_insights = extract_insights("🌍")
    soc_insights = extract_insights("🏢")
    gov_insights = extract_insights("🏛")
    
    return {
        'company_name': company_name,
        'sentiment_score': sentiment_score,
        'environment': env_insights,
        'social': soc_insights,
        'governance': gov_insights
    }

def create_comparison_report():
    """
    Main function to create comparison report from user-specified ESG HTML files
    """
    print("📊 ESG Comparison Report Generator")
    print("="*50)
    print("Please enter the names of the ESG insight HTML files you want to compare (1-5 files)")
    print("Example: ESG_Insights_AAPL.html, ESG_Insights_MSFT.html\n")
    
    while True:
        file_names = input("Enter the HTML filenames (comma separated): ").strip()
        if not file_names:
            print("❌ Please enter at least one filename")
            continue
            
        # Process input
        selected_files = [f.strip() for f in file_names.split(',')]
        valid_files = []
        invalid_files = []
        
        # Validate files
        for f in selected_files[:5]:  # Limit to 5 files
            if not f.endswith('.html'):
                f += '.html'  # Add extension if missing
            
            if os.path.exists(f):
                valid_files.append(f)
            else:
                invalid_files.append(f)
        
        if invalid_files:
            print(f"❌ These files were not found: {', '.join(invalid_files)}")
            if not valid_files:
                continue  # No valid files, ask again
        
        if valid_files:
            break  # We have at least one valid file
    
    # Extract data from each report
    esg_data = []
    for html_file in valid_files:
        try:
            print(f"\n🔍 Processing: {html_file}")
            data = extract_data_from_html(html_file)
            esg_data.append(data)
            print(f"   Company: {data['company_name']}")
            print(f"   ESG Score: {data['sentiment_score']}/10")
            print(f"   Environmental Insights: {len(data['environment'])}")
            print(f"   Social Insights: {len(data['social'])}")
            print(f"   Governance Insights: {len(data['governance'])}")
        except Exception as e:
            print(f"❌ Error processing {html_file}: {str(e)}")
    
    if not esg_data:
        print("❌ No valid ESG data extracted.")
        return
    
    # Generate comparison report
    generate_comparison_html(esg_data)

if __name__ == "__main__":
    create_comparison_report()