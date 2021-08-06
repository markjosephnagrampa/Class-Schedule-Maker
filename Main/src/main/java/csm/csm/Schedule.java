package csm.csm;

public class Schedule {
	String room;
	int day; // 0 - Monday  6 - Sunday
	int[] start, end;
	// 0 - hr
	// 1 - min
	
	public Schedule(int day, int[] start, int[] end, String room) {
		this.day = day;
		this.start = start;
		this.end = end;
		this.room = room;
	}
	/**
	 * gets this schedule's first row number in the formatted Excel schedule file
	 * @return this schedule's first row number
	 */
	public int getStartY() {
		return start[1] == 30 ? start[0] * 2 - 14 + 1 + 3: start[0] * 2 - 14 + 3; 
	}
	/**
	 * gets this schedule's last row number in the formatted Excel schedule file
	 * @return this schedule's last row number
	 */
	public int getEndY() {
		return end[1] == 30 ? end[0] * 2 - 14 + 3: end[0] * 2 - 15 + 3;
	}
	/**
	 * gets this schedule's column number in the formatted Excel schedule file
	 * @return this schedule's cell column number
	 */
	public int getX() {
		return 1 + day;
	}
	/**
	 * checks if this subject will only take a single cell in the formatted Excel schedule file
	 * (a 30 minute class)
	 * @return true if this subject only has a duration of 30 minutes, false otherwise
	 */
	public boolean isSingleBlock() {
		return this.getEndY() - this.getStartY() == 0;
	}
	/**
	 * checks whether 2 subjects conflict with each other's schedule
	 * @param other another Schedule object
	 * @return true if this Schedule conflicts with the Schedule of the other class, false otherwise
	 */
	public boolean isConflicting(Schedule other) {
		
		if(this.day == other.day) {
			double thisStart = this.start[0];
			if(this.start[1] == 30) thisStart += 0.5;
			double thisEnd = this.end[0];
			if(this.end[1] == 30) thisEnd += 0.5;
			
			double otherStart = other.start[0];
			if(other.start[1] == 30) otherStart += 0.5;
			double otherEnd = other.end[0];
			if(other.end[1] == 30) otherEnd += 0.5;
			
			if(thisStart >= otherStart && thisStart < otherEnd)
				return true;
			else if(thisEnd > otherStart && thisEnd <= otherEnd)
				return true;
			else
				return false;
		}
		else {
			return false;
		}
	}
}