#!/usr/bin/env python3
"""
Extract Mathematical Equations from Word Documents

This script extracts mathematical equations from Microsoft Word documents (.docx files).
"""

import os
import sys
import argparse
import re
import logging
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
import mammoth
from bs4 import BeautifulSoup
from docx2python import docx2python

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Define XML namespaces used in DOCX files
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xmlns': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

def extract_equations_docx(file_path, raw_xml=False):
    """
    Extract equations from a .docx file using multiple methods
    
    Args:
        file_path (str): Path to the .docx file
        raw_xml (bool): If True, return the raw XML string without cleaning
        
    Returns:
        list: List of extracted equations
    """
    logger.info(f"Processing DOCX file: {file_path}")
    
    equations = []
    
    try:
        # Use python-docx to read the document
        equations.extend(extract_from_docx_xml(file_path, raw_xml))
        
        # Remove duplicates and empty strings
        equations = [eq.strip() for eq in equations if eq.strip()]
        equations = list(dict.fromkeys(equations))
        
        logger.info(f"Found {len(equations)} equations")
        return equations
    
    except Exception as e:
        logger.error(f"Error processing DOCX file: {e}")
        return []

def extract_from_docx_xml(file_path, raw_xml=False):
    """
    Extract equations by directly parsing the DOCX XML content
    
    Args:
        file_path (str): Path to the .docx file
        raw_xml (bool): If True, return the raw XML string without cleaning
        
    Returns:
        list: List of extracted equations
    """
    logger.debug("Attempting extraction from raw DOCX XML")
    equations = []
    
    try:
        # A .docx file is actually a ZIP archive containing XML files
        with zipfile.ZipFile(file_path) as docx_zip:
            # Check if the document contains math elements
            if 'word/document.xml' in docx_zip.namelist():
                with docx_zip.open('word/document.xml') as xml_file:
                    tree = ET.parse(xml_file)
                    root = tree.getroot()
                    
                    # Find all OMML equation elements
                    # Look for <m:oMath> elements (Office Math Markup Language)
                    for ns in NAMESPACES:
                        # Try different namespaces as they can vary
                        try:
                            math_elements = root.findall(f'.//{{{NAMESPACES[ns]}}}oMath')
                            if math_elements:
                                logger.debug(f"Found {len(math_elements)} math elements with namespace {ns}")
                                
                                for math_elem in math_elements:
                                    # Extract the text content of the math element
                                    equation_text = ET.tostring(math_elem, encoding='unicode')
                                    
                                    if not raw_xml:
                                        # Clean up the XML tags to get just the equation content
                                        equation_text = re.sub(r'<[^>]+>', ' ', equation_text)
                                        equation_text = re.sub(r'\s+', ' ', equation_text).strip()
                                    
                                    if equation_text:
                                        equations.append(equation_text)
                        except Exception as e:
                            logger.debug(f"Error searching namespace {ns}: {e}")
                    
                    # Also look for alternative equation representations
                    # Some equations might be in w:object elements
                    object_elements = root.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}object')
                    for obj in object_elements:
                        obj_text = ET.tostring(obj, encoding='unicode')
                        if 'Equation' in obj_text:
                            if not raw_xml:
                                # Clean up the XML tags
                                obj_text = re.sub(r'<[^>]+>', ' ', obj_text)
                                obj_text = re.sub(r'\s+', ' ', obj_text).strip()
                            
                            if obj_text:
                                equations.append(obj_text)
        
        logger.debug(f"XML extraction found {len(equations)} equations")
        return equations
    
    except Exception as e:
        logger.error(f"Error in XML extraction: {e}")
        return []

def main():
    """Main function to parse arguments and extract equations"""
    parser = argparse.ArgumentParser(description='Extract mathematical equations from Word documents')
    parser.add_argument('file_path', help='Path to the .docx file')
    parser.add_argument('-o', '--output', help='Output file path (default: print to console)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose output')
    parser.add_argument('-r', '--raw', action='store_true', help='Return raw XML without cleaning')
    
    args = parser.parse_args()
    
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    file_path = args.file_path
    raw_xml = args.raw
    
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        sys.exit(1)
    
    # Check file extension
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    if ext == '.docx':
        equations = extract_equations_docx(file_path, raw_xml)
    else:
        logger.error(f"Unsupported file format: {ext}. Please provide a .docx file.")
        sys.exit(1)
    
    # Output results
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            for i, eq in enumerate(equations, 1):
                f.write(f"Equation {i}:\n{eq}\n\n")
        logger.info(f"Results written to {args.output}")
    else:
        print(f"\nExtracted {len(equations)} equations from {file_path}:\n")
        for i, eq in enumerate(equations, 1):
            print(f"Equation {i}:\n{eq}\n")

if __name__ == "__main__":
    main()
