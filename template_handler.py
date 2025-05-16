import os
import re
from string import Formatter
from typing import Dict, List, Set

class TemplateHandler:
    def __init__(self, template_path='email_template.html'):
        """
        Initialize template handler
        
        Args:
            template_path: Path to the HTML template file
        """
        self.template_path = template_path
        self.template_content = self._load_template()
        self.placeholders = self._extract_placeholders()
    
    def _load_template(self) -> str:
        """Load the HTML template from file"""
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template file not found: {self.template_path}")
        
        with open(self.template_path, 'r', encoding='utf-8') as file:
            return file.read()
    
    def _extract_placeholders(self) -> Set[str]:
        """Extract all placeholders ({placeholder}) from the template"""
        # Use string.Formatter to extract all field names from the template
        formatter = Formatter()
        field_names = {
            field_name for _, field_name, _, _ in formatter.parse(self.template_content)
            if field_name is not None
        }
        return field_names
    
    def fill_template(self, data: Dict[str, str]) -> str:
        """
        Fill the template with data
        
        Args:
            data: Dictionary with placeholder keys and their values
            
        Returns:
            Filled HTML template
        """
        # Start with the template
        filled_template = self.template_content
        
        # Replace each placeholder with its corresponding value
        for placeholder in self.placeholders:
            if placeholder in data:
                filled_template = filled_template.replace(
                    '{' + placeholder + '}', 
                    data.get(placeholder, '')
                )
        
        return filled_template
    
    def get_required_fields(self) -> List[str]:
        """Get a list of required fields from the template"""
        return list(self.placeholders)
    
    def reload_template(self):
        """Reload the template from file"""
        self.template_content = self._load_template()
        self.placeholders = self._extract_placeholders() 