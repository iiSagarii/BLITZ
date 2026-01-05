import json
import os
import glob
import re
from dotenv import load_dotenv
from openai import OpenAI
import shutil 

# Load environment variables from .env file
load_dotenv()

# Retrieve the TOKEN
token = os.getenv("TOKEN")
if not token:
    raise ValueError("TOKEN is not set in the .env file")

# Hardcoded paths for files
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ephemeral")  # Points towards the text files generated from the AIP code. [cite: 1]
JSON_OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ephemeral/ai_responses.json")  # AI o/p storage location [cite: 2]
DEBUG_DIR = os.path.join(os.path.dirname(__file__), "ephemeral/debug_outputs")  # Directory for debug outputs [cite: 2]
SYSTEM_MESSAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sys_inst/System_inst-TSS.txt")  # Path to system message file [cite: 2]

# Create debug directory if it doesn't exist
os.makedirs(DEBUG_DIR, exist_ok=True)  

# Initialize OpenAI client with the token
client = OpenAI(  
    base_url="https://openrouter.ai/api/v1",
    api_key=token,
)

def get_system_message():
    """
    Read system message from file.
    
    Returns:
        str: System message content or default message if file not found.
    """
    try:
        if os.path.exists(SYSTEM_MESSAGE_PATH):  
            with open(SYSTEM_MESSAGE_PATH, 'r', encoding='utf-8') as f:  
                return f.read().strip()  
        else:
            print(f"Warning: System message file not found at {SYSTEM_MESSAGE_PATH}")  
            # Default message if file not found
            return "You are a helpful assistant. Always respond with valid JSON only. Do not include any explanatory text, markdown formatting, or comments."  # [cite: 3, 4, 5]
    except Exception as e:
        print(f"Error reading system message file: {str(e)}")  
        # Default message in case of error
        return "You are a helpful assistant. Always respond with valid JSON only. Do not include any explanatory text, markdown formatting, or comments."  # [cite: 4, 5]

def process_file_with_ai(file_path):
    """
    Send the content of a single file to the AI model and return the response. 
    Args:
        file_path (str): Path to the text file to process. 
    Returns:
        str: The AI's response or an error message. 
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:  
            content = f.read()  
        
        # Get system message from file
        system_message = get_system_message()  
        
        response = client.chat.completions.create(  
            messages=[
                {"role": "system", "content": system_message},  
                {"role": "user", "content": content},  
            ],
            model="deepseek/deepseek-r1-distill-qwen-14b",  
            temperature=1,  
            max_tokens=8000,  
            top_p=1  
        )
        
        return response.choices[0].message.content  
    except Exception as e:
        return f"Error processing {file_path}: {str(e)}"  

def fix_json_quotes(json_text):
    """
    Fix improperly terminated quotes in JSON strings that could cause parsing errors. [cite: 10]
    This function specifically addresses:
    1. Double quotes with no content between them ("")
    2. String values ending with a quote character without proper escaping
    
    Args:
        json_text (str): The JSON text to fix
        
    Returns:
        str: Fixed JSON text [cite: 11]
    """
    # Check and fix empty quotes that might cause issues
    json_text = re.sub(r'""', '".\"', json_text)  
    
    # Pattern to find improperly escaped quotes within string values
    # Look for patterns where a string value has an unescaped quote before closing quotes
    def fix_quote_endings(match):  
        content = match.group(1)  
        # If content ends with an unescaped quote, add a period before it
        if content.endswith('"') and not content.endswith('\\"'):  
            return '"' + content[:-1] + '."'  
        return match.group(0)  
    
    # Apply fixes to string values that might end with a quote
    json_text = re.sub(r'"([^"\\]*(?:\\.[^"\\]*)*)"', fix_quote_endings, json_text)  
    
    # Handle specific case where the JSON might be truncated with a hanging quote
    if json_text.endswith('"') and not json_text.endswith('\\"'):  
        json_text = json_text[:-1] + '."'  
    
    # Fix common JSON structural issues
    json_text = json_text.replace('}{', '},{')  
    
    return json_text  

def extract_json(text):
    """
    Attempt to extract valid JSON from a text string using multiple methods. 
    Args:
        text (str): The text potentially containing JSON. 
    Returns:
        tuple: (extracted_json_str, success_flag) 
    """
    # Method 1: Look for JSON in code blocks
    json_block_pattern = r'```(?:json)?(.*?)```'  
    matches = re.findall(json_block_pattern, text, re.DOTALL)  
    if matches:
        return matches[0].strip(), True  
    
    # Method 2: Look for JSON array/object brackets
    try:
        # For arrays
        if '[' in text and ']' in text:  
            start_idx = text.find('[')  
            end_idx = text.rfind(']')  
            if start_idx < end_idx:  
                return text[start_idx:end_idx+1], True  
        
        # For objects
        if '{' in text and '}' in text:  
            start_idx = text.find('{')  
            end_idx = text.rfind('}')  
            if start_idx < end_idx:  
                return text[start_idx:end_idx+1], True  
    except:
        pass  
    
    # No valid JSON structure found
    return text, False  

def parse_json_safely(text):
    """
    Try multiple approaches to parse JSON from text, handling invalid escapes robustly. [cite: 18, 19]

    Args:
        text (str): Text to parse as JSON. [cite: 19]

    Returns:
        tuple: (parsed_data, success_flag) [cite: 20]
    """
    # --- Start Modification ---
    # Clean potential markdown fences and whitespace first from the raw input
    cleaned_text = text.strip()
    if cleaned_text.startswith("```json"):
        # Remove ```json prefix
        cleaned_text = cleaned_text[7:].strip()
        # Remove ``` suffix if present
        if cleaned_text.endswith("```"):
            cleaned_text = cleaned_text[:-3].strip()
    elif cleaned_text.startswith("```"):
        # Remove ``` prefix
        cleaned_text = cleaned_text[3:].strip()
        # Remove ``` suffix if present
        if cleaned_text.endswith("```"):
            cleaned_text = cleaned_text[:-3].strip()
    # Ensure it's stripped again after potential fence removal
    cleaned_text = cleaned_text.strip()

    # Use the cleaned text directly, bypassing the original extract_json call
    # json_text, extracted = extract_json(text) # Commented out original line [cite: 55]
    json_text = cleaned_text # Use the potentially cleaned raw text

    # Apply the JSON quote fix before parsing
    # Pass the cleaned text (which should be the full JSON object) to the fixer
    fixed_json_text = fix_json_quotes(json_text) 
    # --- End Modification ---

    # Save the fixed JSON for debugging (this now reflects fixing the cleaned text)
    debug_fixed_path = os.path.join(DEBUG_DIR, f"fixed_json_{os.urandom(4).hex()}.txt") 
    try:
        with open(debug_fixed_path, 'w', encoding='utf-8') as f: 
            f.write(fixed_json_text) 
    except Exception as e:
        print(f"Warning: Could not write debug file {debug_fixed_path}: {str(e)}")

    # Attempt 1: Parse directly with fixed JSON, fixing invalid escapes iteratively
    max_attempts = 10
    current_json_to_parse = fixed_json_text # Start with the fixed text
    for attempt in range(max_attempts):
        try:
            # Attempt to load the JSON
            return json.loads(current_json_to_parse), True 
        except json.JSONDecodeError as e:
            if "Invalid \\escape" in str(e):
                pos = e.pos
                # Check bounds carefully before attempting fix
                if pos < len(current_json_to_parse):
                     # Try adding a backslash BEFORE the problematic character at e.pos
                     # This handles cases like \", \n etc. needing to become \\", \\n
                     if current_json_to_parse[pos] == '\\' and pos + 1 < len(current_json_to_parse):
                         # If it's already an escape sequence like \n, \t, don't double escape
                         # Check if the character AFTER the backslash is one that forms a valid escape
                         if current_json_to_parse[pos+1] not in ['"', '\\', '/', 'b', 'f', 'n', 'r', 't', 'u']:
                             # It's likely a stray backslash, escape it
                             current_json_to_parse = current_json_to_parse[:pos] + '\\' + current_json_to_parse[pos:]
                             print(f"Attempt {attempt + 1}: Fixed potentially stray backslash at position {pos}")
                         else:
                             # It looks like a valid escape that json.loads just disliked. Stop trying.
                             print(f"Attempt {attempt + 1}: Encountered potentially valid but problematic escape sequence at {pos}. Stopping fix attempts.")
                             break # Exit loop, will proceed to other methods or fail
                     elif current_json_to_parse[pos] != '\\':
                         # If the error is at a character that's NOT a backslash, but is part of an invalid escape
                         # (e.g. "\ ' "), we might not be able to fix it reliably here. Stop trying.
                          print(f"Attempt {attempt + 1}: Encountered invalid escape sequence not starting with '\\' at {pos}. Stopping fix attempts.")
                          break # Exit loop
                     else: # pos is at the end of the string or other edge case
                          break # Exit loop
                else:
                    # Error position is out of bounds, cannot fix
                     print(f"Attempt {attempt + 1}: Invalid escape error position ({pos}) is out of bounds. Stopping fix attempts.")
                     break # Exit loop
            else:
                # Other JSON errors, break the loop and proceed to alternative methods/failure
                print(f"Attempt {attempt + 1}: Encountered non-escape JSONDecodeError: {e}. Stopping iterative fixing.")
                break
        except Exception as general_error: # Catch other potential errors during parsing/fixing
             print(f"Attempt {attempt + 1}: Unexpected error during iterative parsing: {general_error}. Stopping.")
             break # Exit loop

    # If parsing still fails after iterative fixes, log and try alternative methods (if any remain uncommented)
    print(f"Warning: Failed to parse JSON directly after {max_attempts} attempts at fixing invalid escapes. Trying alternative methods if available.") #

    # Attempt 2: Fix common formatting issues (e.g., single object not in array)
    # This attempt uses the *last state* of current_json_to_parse from the loop above
    try:
        if current_json_to_parse.strip().startswith('{') and current_json_to_parse.strip().endswith('}'): # [cite: 61]
            # Try parsing as a single object
            parsed_obj = json.loads(current_json_to_parse) # Try parsing the potentially fixed string
            # Return as a list containing the single object for consistency
            return [parsed_obj], True 
    except json.JSONDecodeError:
         # This method didn't work either
        pass # [cite: 62]
    except Exception as e: # Catch other errors during this attempt
        print(f"Warning: Error during single object parsing attempt: {e}")
        pass

   # Attempt 3: Try to extract individual JSON objects if parsing fails
    """try:
        objects = []  # [cite: 22]
        # Match all top-level JSON objects in the text
        # Use non-recursive regex to find well-formed objects {}
        for match in re.finditer(r'\{(?:[^{}"\']+|\'[^\']*\'|"[^"\\]*(?:\\.[^"\\]*)*")*\}', fixed_json_text):  # [cite: 22]
            try:
                # Apply fix to each potential JSON object
                fixed_obj_text = fix_json_quotes(match.group(0))  
                obj = json.loads(fixed_obj_text)  
                objects.append(obj)  
            except json.JSONDecodeError:
                # Ignore parts that cannot be parsed as individual objects
                continue  
        
        if objects:
            return objects, True  
    except Exception as e:
        # Catch potential regex errors or other issues
        print(f"Warning: Error during individual object extraction: {str(e)}")
        pass  """


    # Failed to parse JSON using any available method
    print(f"Error: Could not parse JSON from response even after cleaning and fixing attempts.") # More specific error
    return None, False # [cite: 65, 24] # Adjusted citation reference if needed
    

def process_and_parse_file(file_path):
    """
    Process a file with AI and parse the JSON response
    Args:
        file_path (str): Path to file to process. 
    Returns:
        tuple: (parsed_data, error_message) 
    """
    file_name = os.path.basename(file_path)  
    
    # Get the AI response
    response = process_file_with_ai(file_path)  
    
    # Save raw response for debugging
    debug_file = os.path.join(DEBUG_DIR, f"raw_{file_name}.txt")  
    try:
        with open(debug_file, 'w', encoding='utf-8') as f:   
            f.write(response)  # 
    except Exception as e:
        print(f"Warning: Could not write debug file {debug_file}: {str(e)}")
        
    # Try to parse JSON
    json_data, success = parse_json_safely(response)  
    
    if success:
        return json_data, None  
    else:
        error_msg = f"Failed to parse valid JSON from response for {file_name}"  
        # Save the problematic response that failed parsing
        failed_parse_file = os.path.join(DEBUG_DIR, f"failed_parse_{file_name}.txt")
        try:
             with open(failed_parse_file, 'w', encoding='utf-8') as f:
                 f.write(response)
        except Exception as e:
            print(f"Warning: Could not write failed parse file {failed_parse_file}: {str(e)}")
        return None, error_msg  

def main():
    """
    Process all user_prompt_TSS-*.txt files and save aggregated AI responses 
    to a single JSON file with "DOC" and "Excel" keys. 
    """
    # Check if the output directory exists
    if not os.path.exists(OUTPUT_DIR):  
        print(f"Error: Output directory not found: {OUTPUT_DIR}")  
        return  
    
    # Find all user_prompt_TSS-*.txt files
    file_pattern = os.path.join(OUTPUT_DIR, "user_prompt_TSS-*.txt")  
    files = glob.glob(file_pattern)  
    if not files:
        print(f"Error: No user_prompt_TSS-*.txt files found in {OUTPUT_DIR}")  
        return  
    
    # Initialize the dictionary to hold aggregated responses
    aggregated_responses = {"DOC": [], "Excel": []}  # Changed from list 
    error_files = []    # List to track files with errors 
    files_with_issues = []  # Track files with missing keys or errors
    
    # Process each file
    for file_path in files:  
        file_name = os.path.basename(file_path)  
        print(f"Processing batch: {file_name}...")  
        
        # Process file and get parsed JSON
        json_data, error = process_and_parse_file(file_path)  
        
        if error:
            # Handle parsing errors
            print(f"Error: {error}")  # [cite: 31]
            error_files.append(file_name)  # [cite: 31]
            files_with_issues.append({"file": file_name, "error": error})  # Add detailed error info
            continue  # Skip aggregation for this file

        if json_data:
            # Ensure json_data is always a list for consistent processing
            if not isinstance(json_data, list):  # [cite: 30]
                json_data = [json_data]  # [cite: 30]
            
            processed_count = 0
            # Aggregate the data
            for item in json_data:  # Iterate through potentially multiple objects in the response
                 if isinstance(item, dict):
                     doc_data = item.get("DOC")  # Use .get() for safer access
                     excel_data = item.get("Excel")

                     # Check if keys exist and data is a list before extending
                     if doc_data is not None and isinstance(doc_data, list):
                         aggregated_responses["DOC"].extend(doc_data)  # Extend the main DOC list
                     elif doc_data is not None:
                          print(f"Warning: 'DOC' data in {file_name} is not a list. Item: {item}")
                          files_with_issues.append({"file": file_name, "warning": "DOC data not a list", "item": item})
                     # else: DOC key missing or None, handled below

                     if excel_data is not None and isinstance(excel_data, list):
                          aggregated_responses["Excel"].extend(excel_data)  # Extend the main Excel list
                     elif excel_data is not None:
                           print(f"Warning: 'Excel' data in {file_name} is not a list. Item: {item}")
                           files_with_issues.append({"file": file_name, "warning": "Excel data not a list", "item": item})
                     # else: Excel key missing or None, handled below
                     
                     # Track if either key was missing
                     if doc_data is None or excel_data is None:
                         missing_keys = []
                         if doc_data is None: missing_keys.append("DOC")
                         if excel_data is None: missing_keys.append("Excel")
                         print(f"Warning: Missing key(s) {missing_keys} in an item from {file_name}. Item: {item}")
                         files_with_issues.append({"file": file_name, "warning": f"Missing key(s): {missing_keys}", "item": item})

                     processed_count += 1  # Count successfully processed items within the file's response
                 else:
                     # Handle cases where an item in the list is not a dictionary
                     print(f"Warning: Found non-dictionary item in response from {file_name}. Item: {item}")
                     files_with_issues.append({"file": file_name, "warning": "Non-dictionary item in response", "item": item})

            if processed_count > 0:  # Only print success if at least one item was processed
                 print(f"Successfully processed and aggregated data from {file_name}")  # [cite: 31] Adjusted success message
            elif not error:  # If no items were processed but there was no parsing error initially
                 print(f"Warning: No valid DOC/Excel data found in the parsed response from {file_name}")
                 # This case might indicate the AI response was valid JSON but not in the expected format
                 files_with_issues.append({"file": file_name, "warning": "Valid JSON but no DOC/Excel data found"})

        else:  # Should not happen if error handling above works, but as a safeguard
             print(f"Error: Unknown issue processing {file_name}")  # [cite: 31]
             error_files.append(file_name)  # [cite: 31]
             files_with_issues.append({"file": file_name, "error": "Unknown processing issue"})  # [cite: 32]
    
    # Save the aggregated responses to the JSON file
    try:
        with open(JSON_OUTPUT_PATH, 'w', encoding='utf-8') as f:  # [cite: 32]
            # Dump the aggregated dictionary, not the list
            json.dump(aggregated_responses, f, indent=4)  # [cite: 32] 
        print(f"Aggregated AI responses saved to {JSON_OUTPUT_PATH}")  # [cite: 32]
        
        if error_files:
            print(f"\n--- Issues Summary ---")
            print(f"Files with parsing errors ({len(error_files)}): {', '.join(error_files)}")  # [cite: 33]
            print(f"Check the {DEBUG_DIR} directory for raw and failed parse responses.")  # [cite: 33]
        
        # Report other issues like missing keys or wrong data types
        if files_with_issues:
             print("\nAdditional processing warnings/errors encountered:")
             for issue in files_with_issues:
                 print(f"  - File: {issue['file']}, Issue: {issue.get('error') or issue.get('warning')}")
                 # Optionally print the problematic item for debugging
                 # if 'item' in issue: 
                 #    print(f"    Item: {issue['item']}")

        # Only clean up files if we had no errors or warnings at all
        if not error_files and not files_with_issues:  # [cite: 33, 34]
            print("\nProcessing completed successfully with no errors or warnings.")  # [cite: 34]
            print("Consider enabling cleanup if desired.")
            # cleanup_files()  # Keep cleanup commented out as per original code [cite: 35]
        else:
            print("\nProcessing completed with some errors or warnings. Keeping input files for debugging.")  # [cite: 35]
            
    except Exception as e:
        print(f"Error saving aggregated JSON file: {str(e)}")  # [cite: 35]

"""
def cleanup_files():
    
    # Delete user prompt files and system message file after successful processing. [cite: 35]
    
    try:
        # Delete all user prompt files
        file_pattern = os.path.join(OUTPUT_DIR, "user_prompt_TSS-*.txt")  # [cite: 36]
        files = glob.glob(file_pattern)  # [cite: 36]
        for file_path in files:  # [cite: 36]
            os.remove(file_path)  # [cite: 36]
            print(f"Deleted: {file_path}")  # [cite: 36]
        
        # Delete system message file
        if os.path.exists(SYSTEM_MESSAGE_PATH):  # [cite: 36]
            os.remove(SYSTEM_MESSAGE_PATH)  # [cite: 37]
            print(f"Deleted: {SYSTEM_MESSAGE_PATH}")  # [cite: 37]
            
        print("Cleanup completed successfully.")  # [cite: 37]
    except Exception as e:
        print(f"Error during cleanup: {str(e)}")  # [cite: 37]
"""

if __name__ == "__main__":
    main()