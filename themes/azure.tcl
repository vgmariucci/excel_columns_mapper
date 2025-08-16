# Azure Theme for TTK
# Based on the Azure design system
# Copyright © 2021 rdbende <rdbende@gmail.com>

package require Tk 8.6

namespace eval ttk::theme::azure-light {
    variable version 1.0
    package provide ttk::theme::azure-light $version

    # Create the azure-light theme
    ttk::style theme create azure-light -parent clam -settings {
        
        # Configure colors
        variable colors
        array set colors {
            -fg             "#000000"
            -bg             "#ffffff" 
            -disabledfg     "#737373"
            -selectfg       "#ffffff"
            -selectbg       "#0078d4"
            -fieldbackground "#ffffff"
            -bordercolor    "#d1d1d1"
            -lightcolor     "#f3f3f3"
            -darkcolor      "#cfcfcf"
            -troughcolor    "#f0f0f0"
            -focuscolor     "#0078d4"
            -checklight     "#ffffff"
        }

        # Configure base styles
        ttk::style configure . \
            -background $colors(-bg) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor) \
            -darkcolor $colors(-darkcolor) \
            -lightcolor $colors(-lightcolor) \
            -troughcolor $colors(-troughcolor) \
            -focuscolor $colors(-focuscolor) \
            -selectbackground $colors(-selectbg) \
            -selectforeground $colors(-selectfg) \
            -fieldbackground $colors(-fieldbackground) \
            -font {-family "Segoe UI" -size 9} \
            -borderwidth 1 \
            -relief flat

        # Button
        ttk::style configure TButton \
            -padding {8 4 8 4} \
            -width -10 \
            -anchor center \
            -foreground $colors(-fg) \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor) \
            -focuscolor $colors(-focuscolor)
            
        ttk::style map TButton \
            -background [list \
                disabled $colors(-lightcolor) \
                pressed $colors(-darkcolor) \
                active $colors(-lightcolor)] \
            -foreground [list \
                disabled $colors(-disabledfg)] \
            -bordercolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -lightcolor [list \
                pressed $colors(-darkcolor)] \
            -darkcolor [list \
                pressed $colors(-lightcolor)]

        # Entry
        ttk::style configure TEntry \
            -padding {8 4 8 4} \
            -fieldbackground $colors(-fieldbackground) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor) \
            -focuscolor $colors(-focuscolor)
            
        ttk::style map TEntry \
            -bordercolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -lightcolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -darkcolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)]

        # Combobox
        ttk::style configure TCombobox \
            -padding {8 4 4 4} \
            -fieldbackground $colors(-fieldbackground) \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -arrowcolor $colors(-fg) \
            -focuscolor $colors(-focuscolor)
            
        ttk::style map TCombobox \
            -bordercolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -arrowcolor [list \
                disabled $colors(-disabledfg)]

        # Frame
        ttk::style configure TFrame \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor)

        # Label
        ttk::style configure TLabel \
            -padding {0 0 0 0} \
            -background $colors(-bg) \
            -foreground $colors(-fg)
            
        ttk::style map TLabel \
            -foreground [list \
                disabled $colors(-disabledfg)]

        # Labelframe
        ttk::style configure TLabelframe \
            -padding {0 0 0 0} \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor)
            
        ttk::style configure TLabelframe.Label \
            -padding {12 6 12 6} \
            -background $colors(-bg) \
            -foreground $colors(-fg)

        # Notebook
        ttk::style configure TNotebook \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -tabmargins {0 2 0 0}
            
        ttk::style configure TNotebook.Tab \
            -padding {12 8 12 8} \
            -background $colors(-lightcolor) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor)
            
        ttk::style map TNotebook.Tab \
            -background [list \
                selected $colors(-bg) \
                active $colors(-lightcolor)] \
            -foreground [list \
                selected $colors(-focuscolor) \
                active $colors(-fg)] \
            -bordercolor [list \
                selected $colors(-focuscolor) \
                active $colors(-bordercolor)]

        # Progressbar
        ttk::style configure TProgressbar \
            -background $colors(-selectbg) \
            -troughcolor $colors(-troughcolor) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-selectbg) \
            -darkcolor $colors(-selectbg)

        # Scrollbar
        ttk::style configure TScrollbar \
            -gripcount 0 \
            -background $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor) \
            -lightcolor $colors(-lightcolor) \
            -troughcolor $colors(-troughcolor) \
            -bordercolor $colors(-bordercolor) \
            -arrowcolor $colors(-fg)
            
        ttk::style map TScrollbar \
            -background [list \
                active $colors(-darkcolor) \
                pressed $colors(-darkcolor)] \
            -arrowcolor [list \
                disabled $colors(-disabledfg)]

        # Separator
        ttk::style configure TSeparator \
            -background $colors(-bordercolor)

        # Treeview
        ttk::style configure Treeview \
            -background $colors(-fieldbackground) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor) \
            -fieldbackground $colors(-fieldbackground)
            
        ttk::style configure Treeview.Heading \
            -padding {8 4 8 4} \
            -background $colors(-lightcolor) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor)
            
        ttk::style map Treeview \
            -background [list \
                selected $colors(-selectbg)] \
            -foreground [list \
                selected $colors(-selectfg)]
            
        ttk::style map Treeview.Heading \
            -background [list \
                active $colors(-darkcolor) \
                pressed $colors(-darkcolor)]
    }
}

namespace eval ttk::theme::azure-dark {
    variable version 1.0
    package provide ttk::theme::azure-dark $version

    # Create the azure-dark theme
    ttk::style theme create azure-dark -parent clam -settings {
        
        # Configure colors
        variable colors
        array set colors {
            -fg             "#ffffff"
            -bg             "#2b2b2b" 
            -disabledfg     "#737373"
            -selectfg       "#ffffff"
            -selectbg       "#0078d4"
            -fieldbackground "#3c3c3c"
            -bordercolor    "#555555"
            -lightcolor     "#404040"
            -darkcolor      "#1a1a1a"
            -troughcolor    "#404040"
            -focuscolor     "#0078d4"
            -checklight     "#ffffff"
        }

        # Configure base styles
        ttk::style configure . \
            -background $colors(-bg) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor) \
            -darkcolor $colors(-darkcolor) \
            -lightcolor $colors(-lightcolor) \
            -troughcolor $colors(-troughcolor) \
            -focuscolor $colors(-focuscolor) \
            -selectbackground $colors(-selectbg) \
            -selectforeground $colors(-selectfg) \
            -fieldbackground $colors(-fieldbackground) \
            -font {-family "Segoe UI" -size 9} \
            -borderwidth 1 \
            -relief flat

        # Button
        ttk::style configure TButton \
            -padding {8 4 8 4} \
            -width -10 \
            -anchor center \
            -foreground $colors(-fg) \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor) \
            -focuscolor $colors(-focuscolor)
            
        ttk::style map TButton \
            -background [list \
                disabled $colors(-lightcolor) \
                pressed $colors(-darkcolor) \
                active $colors(-lightcolor)] \
            -foreground [list \
                disabled $colors(-disabledfg)] \
            -bordercolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -lightcolor [list \
                pressed $colors(-darkcolor)] \
            -darkcolor [list \
                pressed $colors(-lightcolor)]

        # Entry
        ttk::style configure TEntry \
            -padding {8 4 8 4} \
            -fieldbackground $colors(-fieldbackground) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor) \
            -focuscolor $colors(-focuscolor) \
            -insertcolor $colors(-fg)
            
        ttk::style map TEntry \
            -bordercolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -lightcolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -darkcolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)]

        # Combobox
        ttk::style configure TCombobox \
            -padding {8 4 4 4} \
            -fieldbackground $colors(-fieldbackground) \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -arrowcolor $colors(-fg) \
            -focuscolor $colors(-focuscolor)
            
        ttk::style map TCombobox \
            -bordercolor [list \
                focus $colors(-focuscolor) \
                active $colors(-focuscolor)] \
            -arrowcolor [list \
                disabled $colors(-disabledfg)]

        # Frame
        ttk::style configure TFrame \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor)

        # Label
        ttk::style configure TLabel \
            -padding {0 0 0 0} \
            -background $colors(-bg) \
            -foreground $colors(-fg)
            
        ttk::style map TLabel \
            -foreground [list \
                disabled $colors(-disabledfg)]

        # Labelframe
        ttk::style configure TLabelframe \
            -padding {0 0 0 0} \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor)
            
        ttk::style configure TLabelframe.Label \
            -padding {12 6 12 6} \
            -background $colors(-bg) \
            -foreground $colors(-fg)

        # Notebook
        ttk::style configure TNotebook \
            -background $colors(-bg) \
            -bordercolor $colors(-bordercolor) \
            -tabmargins {0 2 0 0}
            
        ttk::style configure TNotebook.Tab \
            -padding {12 8 12 8} \
            -background $colors(-lightcolor) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor)
            
        ttk::style map TNotebook.Tab \
            -background [list \
                selected $colors(-bg) \
                active $colors(-lightcolor)] \
            -foreground [list \
                selected $colors(-focuscolor) \
                active $colors(-fg)] \
            -bordercolor [list \
                selected $colors(-focuscolor) \
                active $colors(-bordercolor)]

        # Progressbar
        ttk::style configure TProgressbar \
            -background $colors(-selectbg) \
            -troughcolor $colors(-troughcolor) \
            -bordercolor $colors(-bordercolor) \
            -lightcolor $colors(-selectbg) \
            -darkcolor $colors(-selectbg)

        # Scrollbar
        ttk::style configure TScrollbar \
            -gripcount 0 \
            -background $colors(-lightcolor) \
            -darkcolor $colors(-darkcolor) \
            -lightcolor $colors(-lightcolor) \
            -troughcolor $colors(-troughcolor) \
            -bordercolor $colors(-bordercolor) \
            -arrowcolor $colors(-fg)
            
        ttk::style map TScrollbar \
            -background [list \
                active $colors(-darkcolor) \
                pressed $colors(-darkcolor)] \
            -arrowcolor [list \
                disabled $colors(-disabledfg)]

        # Separator
        ttk::style configure TSeparator \
            -background $colors(-bordercolor)

        # Treeview
        ttk::style configure Treeview \
            -background $colors(-fieldbackground) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor) \
            -fieldbackground $colors(-fieldbackground)
            
        ttk::style configure Treeview.Heading \
            -padding {8 4 8 4} \
            -background $colors(-lightcolor) \
            -foreground $colors(-fg) \
            -bordercolor $colors(-bordercolor)
            
        ttk::style map Treeview \
            -background [list \
                selected $colors(-selectbg)] \
            -foreground [list \
                selected $colors(-selectfg)]
            
        ttk::style map Treeview.Heading \
            -background [list \
                active $colors(-darkcolor) \
                pressed $colors(-darkcolor)]
    }
}

# Default to light theme
ttk::style theme use azure-light

# Convenience function to switch themes
proc set_theme {mode} {
    if {$mode == "dark"} {
        ttk::style theme use "azure-dark"
    } elseif {$mode == "light"} {
        ttk::style theme use "azure-light"
    }
}