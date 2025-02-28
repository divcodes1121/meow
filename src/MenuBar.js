import React, { useState, useRef, useEffect } from 'react';

// MenuItem Component
const MenuItem = ({ label, shortcut, onClick, disabled, submenu }) => {
  const [showSubmenu, setShowSubmenu] = useState(false);
  const submenuRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (submenuRef.current && !submenuRef.current.contains(event.target)) {
        setShowSubmenu(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);

  return (
    <div 
      style={{
        position: 'relative',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        padding: '8px 16px',
        cursor: disabled ? 'default' : 'pointer',
        color: disabled ? '#aaa' : 'black',
        backgroundColor: showSubmenu ? '#f1f3f4' : 'transparent',
        transition: 'background-color 0.2s ease'
      }}
      onMouseEnter={() => submenu && setShowSubmenu(true)}
      onMouseLeave={() => submenu && setShowSubmenu(false)}
      onClick={!disabled && !submenu ? onClick : undefined}
    >
      <span>{label}</span>
      {shortcut && (
        <span style={{ 
          color: disabled ? '#ccc' : '#888', 
          fontSize: '12px', 
          marginLeft: '16px' 
        }}>
          {shortcut}
        </span>
      )}
      {submenu && showSubmenu && (
        <div 
          ref={submenuRef}
          style={{
            position: 'absolute',
            top: '100%',
            left: 0,
            backgroundColor: 'white',
            border: '1px solid #ddd',
            borderRadius: '4px',
            boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
            zIndex: 1000,
            minWidth: '200px'
          }}
        >
          {submenu.map((subitem, index) => (
            <MenuItem 
              key={index}
              label={subitem.label}
              shortcut={subitem.shortcut}
              onClick={() => {
                setShowSubmenu(false);
                subitem.onClick();
              }}
              disabled={subitem.disabled}
            />
          ))}
        </div>
      )}
    </div>
  );
};

// Dropdown Menu Component
const DropdownMenu = ({ 
  items, 
  isOpen, 
  onClose, 
  position,
  actions
}) => {
  const menuRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (menuRef.current && !menuRef.current.contains(event.target)) {
        onClose();
      }
    };

    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const renderMenuItems = (menuItems) => {
    return menuItems.map((item, index) => {
      // Check if item is a divider
      if (item.type === 'divider') {
        return (
          <div 
            key={`divider-${index}`} 
            style={{ 
              height: '1px', 
              backgroundColor: '#ddd', 
              margin: '4px 0' 
            }} 
          />
        );
      }

      // Regular menu item
      return (
        <MenuItem
          key={item.label}
          label={item.label}
          shortcut={item.shortcut}
          disabled={item.disabled}
          submenu={item.submenu}
          onClick={() => {
            onClose();
            // Check if the item has a specific action
            if (item.action && actions[item.action]) {
              actions[item.action]();
            }
          }}
        />
      );
    });
  };

  return (
    <div 
      ref={menuRef}
      style={{
        position: 'absolute',
        top: position.top,
        left: position.left,
        backgroundColor: 'white',
        border: '1px solid #ddd',
        borderRadius: '4px',
        boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
        zIndex: 1000,
        minWidth: '200px'
      }}
    >
      {renderMenuItems(items)}
    </div>
  );
};

// Main Menu Bar Component
const MenuBar = ({ actions }) => {
  const [openMenu, setOpenMenu] = useState(null);
  const menuBarRef = useRef(null);

  // Menu definitions
  const menus = {
    File: [
      { 
        label: 'New', 
        action: 'newSpreadsheet', 
        shortcut: 'Ctrl+N',
        submenu: [
          { label: 'Blank Workbook', action: 'newSpreadsheet' },
          { label: 'From Template', onClick: () => alert('Template feature coming soon') }
        ]
      },
      { label: 'Open', action: 'openSpreadsheet', shortcut: 'Ctrl+O' },
      { type: 'divider' },
      { 
        label: 'Save', 
        submenu: [
          { label: 'Save', action: 'saveSpreadsheet', shortcut: 'Ctrl+S' },
          { label: 'Save As', onClick: () => alert('Save As feature coming soon') },
          { label: 'Save a Copy', onClick: () => alert('Save a Copy feature coming soon') }
        ]
      },
      { label: 'Export to CSV', action: 'exportToCsv', shortcut: 'Ctrl+E' },
      { type: 'divider' },
      { label: 'Print', action: 'printSpreadsheet', shortcut: 'Ctrl+P' },
      { type: 'divider' },
      { label: 'Close', onClick: () => window.close() }
    ],
    Edit: [
      { label: 'Undo', action: 'undo', shortcut: 'Ctrl+Z' },
      { label: 'Redo', action: 'redo', shortcut: 'Ctrl+Y' },
      { type: 'divider' },
      { label: 'Cut', action: 'cut', shortcut: 'Ctrl+X' },
      { label: 'Copy', action: 'copy', shortcut: 'Ctrl+C' },
      { label: 'Paste', action: 'paste', shortcut: 'Ctrl+V' },
      { type: 'divider' },
      { label: 'Find and Replace', action: 'findReplace', shortcut: 'Ctrl+H' }
    ],
    View: [
      { 
        label: 'Zoom', 
        submenu: [
          { label: 'Zoom In', action: 'zoomIn', shortcut: 'Ctrl+Plus' },
          { label: 'Zoom Out', action: 'zoomOut', shortcut: 'Ctrl+Minus' },
          { label: 'Reset Zoom', onClick: () => alert('Reset Zoom') }
        ]
      },
      { label: 'Full Screen', action: 'toggleFullScreen', shortcut: 'F11' },
      { label: 'Show Formulas', action: 'showFormulas', shortcut: 'Ctrl+`' }
    ],
    Insert: [
      { label: 'Insert Row', action: 'insertRow', shortcut: 'Alt+I R' },
      { label: 'Insert Column', action: 'insertColumn', shortcut: 'Alt+I C' },
      { type: 'divider' },
      { 
        label: 'Functions', 
        submenu: [
          { label: 'Insert Function', action: 'insertFunction', shortcut: 'Shift+F3' },
          { label: 'Financial', onClick: () => alert('Financial Functions') },
          { label: 'Logical', onClick: () => alert('Logical Functions') },
          { label: 'Text', onClick: () => alert('Text Functions') }
        ]
      },
      { label: 'Chart', action: 'insertChart', shortcut: 'Alt+F1' }
    ],
    Format: [
      { 
        label: 'Number', 
        submenu: [
          { label: 'Number Format', action: 'numberFormat', shortcut: 'Ctrl+Shift+1' },
          { label: 'Percentage', onClick: () => alert('Percentage Format') },
          { label: 'Currency', onClick: () => alert('Currency Format') }
        ]
      },
      { label: 'Bold', action: 'toggleBold', shortcut: 'Ctrl+B' },
      { label: 'Italic', action: 'toggleItalic', shortcut: 'Ctrl+I' },
      { label: 'Underline', action: 'toggleUnderline', shortcut: 'Ctrl+U' },
      { type: 'divider' },
      { label: 'Merge Cells', action: 'mergeCells', shortcut: 'Alt+H M' }
    ],
    Data: [
      { label: 'Sort', action: 'sortData', shortcut: 'Alt+H S' },
      { label: 'Filter', action: 'filterData', shortcut: 'Ctrl+Shift+L' },
      { type: 'divider' },
      { label: 'Data Validation', action: 'dataValidation', shortcut: 'Alt+D V' },
      { label: 'Remove Duplicates', action: 'removeDuplicates', shortcut: 'Alt+D D' }
    ],
    Tools: [
      { label: 'Spelling', action: 'checkSpelling', shortcut: 'F7' },
      { label: 'Word Count', action: 'wordCount', shortcut: 'Ctrl+Shift+C' },
      { type: 'divider' },
      { label: 'Protect Sheet', action: 'protectSheet', shortcut: 'Alt+T P' }
    ],
    Help: [
      { label: 'Help', action: 'showHelp', shortcut: 'F1' },
      { label: 'About', action: 'showAbout', shortcut: '' }
    ]
  };

  const handleMenuToggle = (menuName) => {
    setOpenMenu(openMenu === menuName ? null : menuName);
  };

  const handleMenuClose = () => {
    setOpenMenu(null);
  };

  return (
    <div 
      ref={menuBarRef}
      style={{ 
        display: 'flex', 
        padding: '8px 16px',
        borderBottom: '1px solid #ccc',
        backgroundColor: '#f8f9fa',
        position: 'relative'
      }}
    >
      {Object.keys(menus).map((menuName) => (
        <div 
          key={menuName}
          style={{ 
            position: 'relative',
            marginRight: '16px',
            cursor: 'pointer'
          }}
        >
          <div 
            onClick={() => handleMenuToggle(menuName)}
            style={{
              padding: '4px 8px',
              borderRadius: '4px',
              backgroundColor: openMenu === menuName ? '#e6f4ea' : 'transparent',
              transition: 'background-color 0.2s ease'
            }}
          >
            {menuName}
          </div>
          {openMenu === menuName && (
            <DropdownMenu
              items={menus[menuName]}
              isOpen={true}
              onClose={handleMenuClose}
              position={{
                top: '100%',
                left: 0
              }}
              actions={actions}
            />
          )}
        </div>
      ))}
    </div>
  );
};

export default MenuBar;