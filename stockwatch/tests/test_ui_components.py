"""Tests for UI components."""

import unittest
from unittest.mock import MagicMock, patch
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtWidgets import QApplication, QLabel

import sys
from stockwatch.views.ui_components import ClickableLabel

# Create QApplication instance for testing Qt components
app = QApplication.instance()
if not app:
    app = QApplication(sys.argv)

class TestUIComponents(unittest.TestCase):
    """Test case for UI components."""
    
    def test_clickable_label_init(self):
        """Test ClickableLabel initialization."""
        test_text = "Test Label"
        test_url = "https://www.example.com"
        label = ClickableLabel(test_text, test_url)
        
        # Check that the label has the correct text
        self.assertEqual(label.text(), test_text)
        
        # Check that the label stores the URL
        self.assertEqual(label.url, test_url)
        
        # Check that the cursor is a pointing hand
        self.assertEqual(label.cursor().shape(), Qt.PointingHandCursor)
        
        # Check styling
        self.assertIn("color: blue", label.styleSheet())
        self.assertIn("text-decoration: underline", label.styleSheet())
    
    @patch('stockwatch.views.ui_components.QDesktopServices.openUrl')
    def test_clickable_label_mouse_press(self, mock_open_url):
        """Test that clicking the label opens the URL."""
        test_url = "https://www.example.com"
        label = ClickableLabel("Test", test_url)
        
        # Create a mock mouse event
        mock_event = MagicMock()
        mock_event.button.return_value = Qt.LeftButton
        
        # Simulate clicking the label
        label.mousePressEvent(mock_event)
        
        # Check that openUrl was called with the correct URL
        mock_open_url.assert_called_once()
        url_arg = mock_open_url.call_args[0][0]
        self.assertEqual(url_arg.toString(), test_url)
        
    def test_clickable_label_right_click(self):
        """Test that right-clicking the label doesn't open the URL."""
        with patch('stockwatch.views.ui_components.QDesktopServices.openUrl') as mock_open_url:
            test_url = "https://www.example.com"
            label = ClickableLabel("Test", test_url)
            
            # Create a mock mouse event for right click
            mock_event = MagicMock()
            mock_event.button.return_value = Qt.RightButton
            
            # Simulate right-clicking the label
            label.mousePressEvent(mock_event)
            
            # Check that openUrl was not called
            mock_open_url.assert_not_called()

if __name__ == '__main__':
    unittest.main()