#!/usr/bin/env python3
"""
CV FMV Calculator - Production Level Script
Calculates Fair Market Value (FMV) for all doctor entries from CVdump.csv
Based on scoring_criteria.xlsx and OUS FMV Rates
Author: AI Assistant
Version: 2.0 - Corrected for 100% accuracy matching FMVcalnew.py
"""

import pandas as pd
import os
import sys
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import traceback

# =============================================================================
# CONFIGURATION & LOGGING SETUP
# =============================================================================

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('cv_fmv_calculator.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# File paths
CVDUMP_FILE = "CVdump.csv"
SCORING_CRITERIA_FILE = "scoring_criteria.xlsx"
OUTPUT_FILE = f"CV_FMV_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# =============================================================================
# SCORING CRITERIA LOOKUP DICTIONARIES
# =============================================================================

def create_scoring_lookup():
    """Create comprehensive scoring lookup dictionaries"""
    
    # Years of Experience scoring
    years_experience_scores = {
        "1-2 years of experience": 0,
        "3-7 years of experience": 2,
        "8-14 years of experience": 4,
        "15+ years of experience": 6
    }
    
    # Clinical Experience scoring
    clinical_experience_scores = {
        "Minimal patient interactions and predominantly administrative/academic work": 0,
        "Less than half the time spent with patients in clinical setting and higher focus on academic/administrative work": 2,
        "Equal amount of time spent with patients in clinical setting and equal amount of time spent in academic/administrative work": 4,
        "Significant time spent with patients in clinical setting and minimal time spent in academic/administrative work": 6
    }
    
    # Leadership position scoring
    leadership_scores = {
        "Not applicable, as not a part of any society or leadership roles in hospital": 0,
        "1-2 years in a leadership position(s) eg. HOD of a particular speciality in Hospital or other Patient Care Setting and/or serving as a President, Vice president, Secretary,Treasurer, Board member in a Professional or Scientific Society.": 2,
        "3-7 years in a leadership position(s) eg HOD of a particular speciality   in Hospital or other Patient Care Setting and/or serving as a national/regional leader in a Professional or Scientific Society.": 4,
        "8 or more years in a leadership position(s) eg HOD for a specialty in Hospital or other Patient Care Setting and/or serving as an international leader in a Professional or Scientific Society.": 6
    }
    
    # Geographical Reach scoring
    geographical_reach_scores = {
        "Local Influence": 0,
        "National Influence": 2,
        "Multi-Country Influence": 4,
        "Global/Worldwide Influence": 6
    }
    
    # Highest Academic Position scoring
    academic_position_scores = {
        "None or N/A": 0,
        "Professor (including Associate / Assistant Professor)": 2,
        "Professor or Adjunct/Additional/Emeritus Professor": 4,
        "Department Chair/ HOD (or similar position)": 6
    }
    
    # Additional Educational Level scoring
    educational_level_scores = {
        "None or N/A": 0,
        "1 Additional degree, fellowship, or advanced training certification.": 2,
        "2 Additional degrees, fellowship, or advanced training certification.": 4,
        "3 or More Additional degrees, fellowship, or advanced training certification.": 6
    }
    
    # Research Experience scoring
    research_experience_scores = {
        "None or N/A": 0,
        "Participation as an Investigator or Sub-Investigator in 1 to 4 clinical trials or research studies.": 2,
        "Participation as an Investigator or Sub-Investigator in 5 to 9 clinical trials or research studies.": 4,
        "Participation as an Investigator of Sub-Investigator in 10 or more clinical trials or research studies or Principal Investigator for two or more clinical trials or research studies or serving as the Principal Investigator for a clinical trial or research study that led to important medical innovations or significant medical technology breakthroughs.": 6
    }
    
    # Publication Experience scoring
    publication_experience_scores = {
        "None or N/A": 0,
        "Co-authorship or participation as contributing author on 1 to 4 publications.": 2,
        "First authorship (if known) on 1 to 5 publications and/or co-authorship or participation as contributing author on 6 to 10 publications": 4,
        "First authorship (if known) on 6 or more publications and/or co-authorship or participation as contributing author on 11 or more publications": 6
    }
    
    # Speaking Experience scoring
    speaking_experience_scores = {
        "Local speaking engagements and the scientific work done for the specialty is near to the practice location": 0,
        "Most of the speaking engagements are directed nationally for the conferences, symposia or national webinars in the designated specialty and the scientific work done is not restricted for the local audience": 2,
        "The speaking experiences are not restricted nationally but to a group of specified countries and the scientific work is directed to the same group of countries": 4,
        "The speaking engagements and the scinetific work carried out is across the globe": 6
    }
    
    return {
        "years_experience": years_experience_scores,
        "clinical_experience": clinical_experience_scores,
        "leadership": leadership_scores,
        "geographical_reach": geographical_reach_scores,
        "academic_position": academic_position_scores,
        "educational_level": educational_level_scores,
        "research_experience": research_experience_scores,
        "publication_experience": publication_experience_scores,
        "speaking_experience": speaking_experience_scores
    }

# =============================================================================
# FMV RATES LOADING
# =============================================================================

def load_fmv_rates():
    """Load FMV rates from OUS FMV Rates sheet - returns DataFrame for better matching"""
    try:
        rates_df = pd.read_excel(SCORING_CRITERIA_FILE, sheet_name="OUS FMV Rates", header=1)
        
        # Filter for India rates
        india_rates = rates_df[rates_df['Country'] == 'India'].copy()
        
        logger.info(f"Loaded FMV rates for {len(india_rates)} India specialty entries")
        return india_rates
    except Exception as e:
        logger.error(f"Error loading FMV rates: {str(e)}")
        raise

# =============================================================================
# SCORING FUNCTIONS
# =============================================================================

def find_column_name(df, possible_names):
    """Find the actual column name from a list of possible names"""
    for name in possible_names:
        if name in df.columns:
            return name
    # Try case-insensitive match
    for name in possible_names:
        for col in df.columns:
            if col.lower() == name.lower():
                return col
    # Try partial match
    for name in possible_names:
        for col in df.columns:
            if name.lower() in col.lower() or col.lower() in name.lower():
                return col
    return None

def safe_get_value(row, col_name):
    """Safely get value from row, handling NaN and empty strings"""
    if not col_name or col_name not in row.index:
        return ""
    value = row[col_name]
    if pd.isna(value) or str(value).lower() == "nan":
        return ""
    return str(value).strip()

def calculate_individual_scores(row, scoring_lookup, column_mapping):
    """Calculate individual scores for each criterion with proper column name handling"""
    scores = {}
    
    # Years of Experience (Score 1) - Handle column name variations
    years_col = column_mapping.get("years_experience_col")
    years_exp = safe_get_value(row, years_col)
    scores["score_1"] = scoring_lookup["years_experience"].get(years_exp, 0)
    
    # Clinical Experience (Score 2)
    clinical_col = column_mapping.get("clinical_experience_col")
    clinical_exp = safe_get_value(row, clinical_col)
    scores["score_2"] = scoring_lookup["clinical_experience"].get(clinical_exp, 0)
    
    # Leadership position (Score 3)
    leadership_col = column_mapping.get("leadership_col")
    leadership = safe_get_value(row, leadership_col)
    scores["score_3"] = scoring_lookup["leadership"].get(leadership, 0)
    
    # Geographical Reach (Score 4)
    geo_col = column_mapping.get("geographical_reach_col")
    geo_reach = safe_get_value(row, geo_col)
    scores["score_4"] = scoring_lookup["geographical_reach"].get(geo_reach, 0)
    
    # Highest Academic Position (Score 5)
    academic_col = column_mapping.get("academic_position_col")
    academic_pos = safe_get_value(row, academic_col)
    scores["score_5"] = scoring_lookup["academic_position"].get(academic_pos, 0)
    
    # Additional Educational Level (Score 6)
    edu_col = column_mapping.get("educational_level_col")
    edu_level = safe_get_value(row, edu_col)
    scores["score_6"] = scoring_lookup["educational_level"].get(edu_level, 0)
    
    # Research Experience (Score 7)
    research_col = column_mapping.get("research_experience_col")
    research_exp = safe_get_value(row, research_col)
    scores["score_7"] = scoring_lookup["research_experience"].get(research_exp, 0)
    
    # Publication Experience (Score 8)
    pub_col = column_mapping.get("publication_experience_col")
    pub_exp = safe_get_value(row, pub_col)
    scores["score_8"] = scoring_lookup["publication_experience"].get(pub_exp, 0)
    
    # Speaking Experience (Score 9)
    speaking_col = column_mapping.get("speaking_experience_col")
    speaking_exp = safe_get_value(row, speaking_col)
    scores["score_9"] = scoring_lookup["speaking_experience"].get(speaking_exp, 0)
    
    # Calculate total score (excluding total_score from the sum)
    scores["total_score"] = sum([scores["score_1"], scores["score_2"], scores["score_3"], 
                                 scores["score_4"], scores["score_5"], scores["score_6"],
                                 scores["score_7"], scores["score_8"], scores["score_9"]])
    
    return scores

def determine_tier(total_score):
    """Determine tier based on total score"""
    if total_score <= 13:
        return "Tier 1"
    elif total_score <= 26:
        return "Tier 2"
    elif total_score <= 40:
        return "Tier 3"
    else:
        return "Tier 4"

def calculate_fmv_amount(specialty, tier, rates_df):
    """Calculate FMV amount based on specialty and tier using precise matching like FMVcalnew.py"""
    try:
        # Clean specialty name for better matching
        specialty_clean = str(specialty).strip()
        
        if pd.isna(specialty_clean) or specialty_clean == "" or specialty_clean == "nan":
            logger.warning(f"Empty specialty for tier {tier}, using default")
            default_rates = {
                "Tier 1": 5000,
                "Tier 2": 7000,
                "Tier 3": 9000,
                "Tier 4": 12000
            }
            return default_rates.get(tier, 5000)
        
        # Find matching specialty in rates (exact match first)
        specialty_row = rates_df[rates_df["HCP Specialty"] == specialty_clean]
        
        if specialty_row.empty:
            # Try case-insensitive exact match
            specialty_row = rates_df[rates_df["HCP Specialty"].str.lower() == specialty_clean.lower()]
        
        if specialty_row.empty:
            # Try partial matching for specialties that might have slight variations
            specialty_row = rates_df[rates_df["HCP Specialty"].str.contains(specialty_clean, case=False, na=False)]
        
        if specialty_row.empty:
            # Log the specialty that wasn't found for debugging
            logger.warning(f"Specialty not found in rates table: '{specialty_clean}'")
            # Use a conservative default rate based on tier
            default_rates = {
                "Tier 1": 5000,
                "Tier 2": 7000,
                "Tier 3": 9000,
                "Tier 4": 12000
            }
            return default_rates.get(tier, 5000)
        
        # Get the rate for the tier - use exact column names with spaces
        if tier in specialty_row.columns:
            rate = specialty_row[tier].iloc[0]
            # Ensure we return a whole number, handle both int and float
            if pd.isna(rate):
                return 0
            try:
                return int(float(rate))
            except (ValueError, TypeError):
                return 0
        else:
            # Log the tier that wasn't found
            logger.warning(f"Tier column not found: '{tier}' in columns: {list(specialty_row.columns)}")
            default_rates = {
                "Tier 1": 5000,
                "Tier 2": 7000,
                "Tier 3": 9000,
                "Tier 4": 12000
            }
            return default_rates.get(tier, 5000)
            
    except Exception as e:
        logger.warning(f"Error calculating honorarium rate for '{specialty}', '{tier}': {str(e)}")
        default_rates = {
            "Tier 1": 5000,
            "Tier 2": 7000,
            "Tier 3": 9000,
            "Tier 4": 12000
        }
        return default_rates.get(tier, 5000)

# =============================================================================
# DATA PROCESSING FUNCTIONS
# =============================================================================

def load_cvdump_data():
    """Load and clean CVdump.csv data with proper column name detection"""
    try:
        logger.info("Loading CVdump.csv data...")
        
        # Try different encodings
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(CVDUMP_FILE, encoding=encoding, dtype=str)
                logger.info(f"Successfully loaded CVdump.csv with {encoding} encoding")
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception as e:
                logger.warning(f"Error with {encoding} encoding: {str(e)}")
                continue
        
        if df is None:
            raise Exception("Could not load CVdump.csv with any supported encoding")
        
        # Clean email addresses
        if "HCP Email" in df.columns:
            df["HCP Email"] = df["HCP Email"].astype(str).str.strip().str.lower()
            # Remove rows with invalid emails
            df = df[df["HCP Email"] != "nan"]
            df = df[df["HCP Email"] != ""]
        else:
            logger.error("HCP Email column not found in CVdump.csv")
            raise ValueError("HCP Email column not found")
        
        logger.info(f"Loaded {len(df)} records from CVdump.csv")
        logger.info(f"Columns found: {list(df.columns)[:5]}...")  # Log first 5 columns
        
        return df
    except Exception as e:
        logger.error(f"Error loading CVdump data: {str(e)}")
        logger.error(traceback.format_exc())
        raise

def detect_column_names(df):
    """Detect actual column names in the DataFrame and create mapping"""
    column_mapping = {}
    
    # Years of experience - try multiple variations
    years_variations = [
        "Years of experience in the\xa0Specialty / Super Specialty?\n",  # With non-breaking space
        "Years of experience in the Specialty / Super Specialty?\n",
        "Years of experience in the Specialty / Super Specialty?",
        "Years of experience in the\xa0Specialty / Super Specialty?"
    ]
    years_col = find_column_name(df, years_variations)
    if years_col:
        column_mapping["years_experience_col"] = years_col
        logger.info(f"Found years column: '{years_col}'")
    else:
        logger.warning("Could not find years of experience column")
        column_mapping["years_experience_col"] = None
    
    # Clinical Experience
    clinical_variations = [
        "Clinical Experience: i.e. Time Spent with Patients?",
        "Clinical Experience"
    ]
    clinical_col = find_column_name(df, clinical_variations)
    if clinical_col:
        column_mapping["clinical_experience_col"] = clinical_col
    else:
        logger.warning("Could not find clinical experience column")
        column_mapping["clinical_experience_col"] = None
    
    # Leadership
    leadership_variations = [
        "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...",
        "Leadership position"
    ]
    leadership_col = find_column_name(df, leadership_variations)
    if leadership_col:
        column_mapping["leadership_col"] = leadership_col
    else:
        logger.warning("Could not find leadership column")
        column_mapping["leadership_col"] = None
    
    # Geographical Reach
    geo_variations = [
        "Geographic influence as a Key Opinion Leader.",
        "Geographic influence",
        "Geographical Reach"
    ]
    geo_col = find_column_name(df, geo_variations)
    if geo_col:
        column_mapping["geographical_reach_col"] = geo_col
    else:
        logger.warning("Could not find geographical reach column")
        column_mapping["geographical_reach_col"] = None
    
    # Academic Position
    academic_variations = [
        "Highest Academic Position Held in past 10 years",
        "Highest Academic Position"
    ]
    academic_col = find_column_name(df, academic_variations)
    if academic_col:
        column_mapping["academic_position_col"] = academic_col
    else:
        logger.warning("Could not find academic position column")
        column_mapping["academic_position_col"] = None
    
    # Educational Level
    edu_variations = [
        "Additional Educational Level ",
        "Additional Educational Level",
        "Additional Education"
    ]
    edu_col = find_column_name(df, edu_variations)
    if edu_col:
        column_mapping["educational_level_col"] = edu_col
    else:
        logger.warning("Could not find educational level column")
        column_mapping["educational_level_col"] = None
    
    # Research Experience
    research_variations = [
        "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years",
        "Research Experience"
    ]
    research_col = find_column_name(df, research_variations)
    if research_col:
        column_mapping["research_experience_col"] = research_col
    else:
        logger.warning("Could not find research experience column")
        column_mapping["research_experience_col"] = None
    
    # Publication Experience
    pub_variations = [
        "Publication experience in the past 10 years",
        "Publication experience",
        "Publication Experience"
    ]
    pub_col = find_column_name(df, pub_variations)
    if pub_col:
        column_mapping["publication_experience_col"] = pub_col
    else:
        logger.warning("Could not find publication experience column")
        column_mapping["publication_experience_col"] = None
    
    # Speaking Experience
    speaking_variations = [
        "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.",
        "Speaking experience",
        "Speaking Experience"
    ]
    speaking_col = find_column_name(df, speaking_variations)
    if speaking_col:
        column_mapping["speaking_experience_col"] = speaking_col
    else:
        logger.warning("Could not find speaking experience column")
        column_mapping["speaking_experience_col"] = None
    
    # Other important columns
    if "HCP Name" in df.columns:
        column_mapping["hcp_name_col"] = "HCP Name"
    if "HCP Email" in df.columns:
        column_mapping["hcp_email_col"] = "HCP Email"
    if "Specialty / Super Specialty" in df.columns:
        column_mapping["specialty_col"] = "Specialty / Super Specialty"
    if "Educational Qualification" in df.columns:
        column_mapping["qualification_col"] = "Educational Qualification"
    
    return column_mapping

def process_doctor_data(df, scoring_lookup, rates_df, column_mapping):
    """Process each doctor's data and calculate FMV"""
    results = []
    
    for index, row in df.iterrows():
        try:
            # Calculate individual scores
            scores = calculate_individual_scores(row, scoring_lookup, column_mapping)
            total_score = scores["total_score"]
            
            # Determine tier
            tier = determine_tier(total_score)
            
            # Get specialty
            specialty_col = column_mapping.get("specialty_col", "Specialty / Super Specialty")
            specialty = safe_get_value(row, specialty_col)
            
            # Calculate FMV amount using DataFrame-based matching
            fmv_amount = calculate_fmv_amount(specialty, tier, rates_df)
            
            # Get column names for output
            years_col = column_mapping.get("years_experience_col", "Years of experience in the Specialty / Super Specialty?\n")
            hcp_name_col = column_mapping.get("hcp_name_col", "HCP Name")
            hcp_email_col = column_mapping.get("hcp_email_col", "HCP Email")
            qualification_col = column_mapping.get("qualification_col", "Educational Qualification")
            
            # Create result record matching FMV_Calculator_Updated.xlsx structure
            result = {
                "i": index + 1,  # Sequential number
                "HCP Name": row.get(hcp_name_col, "") if hcp_name_col else "",
                "Years of experience in the Specialty / Super Specialty?_x000D_\n": row.get(years_col, "") if years_col else "",
                "Clinical Experience: i.e. Time Spent with Patients?": row.get(column_mapping.get("clinical_experience_col", ""), "") if column_mapping.get("clinical_experience_col") else "",
                "Leadership position(s) in a Professional or Scientific Society and/or leadership position(s) in Hospital or other Patient Care Settings (e.g. Department Head or Chief, Medical Director, Lab Direct...": row.get(column_mapping.get("leadership_col", ""), "") if column_mapping.get("leadership_col") else "",
                "Geographic influence as a Key Opinion Leader.": row.get(column_mapping.get("geographical_reach_col", ""), "") if column_mapping.get("geographical_reach_col") else "",
                "Highest Academic Position Held in past 10 years": row.get(column_mapping.get("academic_position_col", ""), "") if column_mapping.get("academic_position_col") else "",
                "Additional Educational Level": row.get(column_mapping.get("educational_level_col", ""), "") if column_mapping.get("educational_level_col") else "",
                "Research Experience (e.g., industry-sponsored research, investigator-initiated research, other research) in past 10 years": row.get(column_mapping.get("research_experience_col", ""), "") if column_mapping.get("research_experience_col") else "",
                "Publication experience in the past 10 years": row.get(column_mapping.get("publication_experience_col", ""), "") if column_mapping.get("publication_experience_col") else "",
                "Speaking experience (professional, academic, scientific, or media experience) in the past 10 years.": row.get(column_mapping.get("speaking_experience_col", ""), "") if column_mapping.get("speaking_experience_col") else "",
                "Score based on selection mentioned criteria": total_score,
                "Score 1": scores["score_1"],
                "Score 2": scores["score_2"],
                "Score 3": scores["score_3"],
                "Score 4": scores["score_4"],
                "Score 5": scores["score_5"],
                "Score 6": scores["score_6"],
                "Score 7": scores["score_7"],
                "Score 8": scores["score_8"],
                "Score 9": scores["score_9"],
                "Range": f"{total_score}-{total_score}",  # Individual score range
                "Tier": tier,
                "Rate of Honorarium": fmv_amount,
                "Specialty / Super Specialty": specialty,
                "HCP Email": row.get(hcp_email_col, "") if hcp_email_col else "",
                "Educational Qualification": row.get(qualification_col, "") if qualification_col else ""
            }
            
            results.append(result)
            
        except Exception as e:
            logger.error(f"Error processing doctor {row.get('HCP Name', 'Unknown')} (row {index + 1}): {str(e)}")
            logger.error(traceback.format_exc())
            continue
    
    return results

def save_results(results):
    """Save results to Excel file"""
    try:
        results_df = pd.DataFrame(results)
        
        # Create Excel file with single sheet matching FMV_Calculator_Updated.xlsx structure
        results_df.to_excel(OUTPUT_FILE, sheet_name='Sheet1', index=False, engine='openpyxl')
        
        logger.info(f"Results saved to {OUTPUT_FILE}")
        return OUTPUT_FILE
        
    except Exception as e:
        logger.error(f"Error saving results: {str(e)}")
        logger.error(traceback.format_exc())
        raise

# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main execution function"""
    try:
        logger.info("=" * 60)
        logger.info("STARTING CV FMV CALCULATOR")
        logger.info("=" * 60)
        
        # Load scoring criteria
        logger.info("Loading scoring criteria...")
        scoring_lookup = create_scoring_lookup()
        
        # Load FMV rates (returns DataFrame for better matching)
        logger.info("Loading FMV rates...")
        rates_df = load_fmv_rates()
        
        # Load CVdump data
        logger.info("Loading CVdump data...")
        cvdump_df = load_cvdump_data()
        
        # Detect column names
        logger.info("Detecting column names...")
        column_mapping = detect_column_names(cvdump_df)
        
        # Process doctor data
        logger.info("Processing doctor data and calculating FMV...")
        results = process_doctor_data(cvdump_df, scoring_lookup, rates_df, column_mapping)
        
        # Save results
        logger.info("Saving results...")
        output_file = save_results(results)
        
        logger.info("=" * 60)
        logger.info("✅ CV FMV CALCULATOR COMPLETED SUCCESSFULLY")
        logger.info("=" * 60)
        logger.info(f"📊 Processed {len(results)} doctors")
        logger.info(f"📁 Results saved to: {output_file}")
        
        # Print summary
        if results:
            total_fmv = sum(r['Rate of Honorarium'] for r in results)
            avg_score = sum(r['Score based on selection mentioned criteria'] for r in results) / len(results)
            tier_counts = {}
            for r in results:
                tier = r['Tier']
                tier_counts[tier] = tier_counts.get(tier, 0) + 1
            
            print(f"\n📈 SUMMARY:")
            print(f"   Total Doctors: {len(results)}")
            print(f"   Average Score: {avg_score:.2f}")
            print(f"   Total FMV Amount: ₹{total_fmv:,}")
            print(f"   Tier Distribution:")
            for tier, count in sorted(tier_counts.items()):
                print(f"     {tier}: {count} doctors")
        
    except Exception as e:
        logger.error("=" * 60)
        logger.error("❌ ERROR IN MAIN EXECUTION")
        logger.error("=" * 60)
        logger.error(f"Error: {str(e)}")
        logger.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()
