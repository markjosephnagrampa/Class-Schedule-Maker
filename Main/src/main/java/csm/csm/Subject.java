package csm.csm;

import java.awt.Color;
import java.util.ArrayList;

public class Subject {
	
	static Color[] colorPallete = {
			new Color(141,180,227), // Blue
			new Color(255,153,204), // Pink
			new Color(250,192,144), // Orange
			new Color(204,192,218), // Purple
			new Color(255,255,102), // Yellow
			new Color(191,191,191), // Gray
			new Color(182,221,232), // Sky Blue
			new Color(229,184,183), // Carnation Pink
			new Color(146,208,80), // Dark Green
			new Color(214, 227, 188), // Olive Green
			new Color(204,153,0), // Gold
			new Color(223,213,165), // Faint Brown
			new Color(94,43,255), // Royal Blue
			new Color(255,255,255) // White
			
			};
	String name;
	Color color;
	ArrayList<Schedule> schedules = new ArrayList<Schedule>();
	public Subject(String name) {
		this.name = name;
	}
	/**
	 * returns this subject's formatted name (abbreviated if it exceeds the cell width)
	 * @return the formatted subject name
	 */
	public String getFormattedName() {
		if(name.length() <= 12)
			return name;
		else {
			String[] text = name.split(" ");
			String abbrText = "";
			for(String s: text) {
				if(Character.isUpperCase(s.charAt(0)))
					abbrText += s.charAt(0);
				else {
					String numberPattern = "\\d+";
					if(s.matches(numberPattern))
						abbrText += " " + s;
				}
				
			}
			
			if(abbrText.length() <= 12)
				return abbrText;
			else
				return abbrText.substring(0,9) + "...";
		}
	}
}