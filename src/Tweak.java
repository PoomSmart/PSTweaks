import java.util.Map.Entry;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Tweak {
	public String name;
	public Version version;
	public boolean supports64Bit = true;
	public boolean available = true;
	public boolean fromMyRepo;
	public boolean isFlipswitch;
	public String comments;
	public TreeMap<Integer, Support> supports;

	public static Pattern iosVerPattern = Pattern.compile("([l\\-\\d+]+):(.+)");

	public Tweak(String name, boolean supports64Bit, boolean available, boolean fromMyRepo, String comments) {
		this.name = name;
		this.supports64Bit = supports64Bit;
		this.available = available;
		this.fromMyRepo = fromMyRepo;
		this.comments = comments;
		this.supports = new TreeMap<Integer, Support>();
	}

	public Tweak(String name, boolean fromMyRepo, String comments) {
		this(name, true, true, fromMyRepo, comments);
	}

	public Tweak(String desc) {
		// example: 60fps,s64,my,[comment]
		assert desc != null && !desc.isEmpty();
		this.supports = new TreeMap<Integer, Support>();
		int open = desc.lastIndexOf('[');
		int close = desc.lastIndexOf(']');
		if (open != -1 && close != -1 && open + 1 < close) {
			this.comments = desc.substring(open + 1, close);
			desc = desc.replace('[' + this.comments + ']', "");
		}
		String[] sub = desc.split(",");
		this.name = sub[0];
		for (int i = 1; i < sub.length; i++) {
			String s = sub[i];
			switch (s) {
			case "n64":
				supports64Bit = false;
				break;
			case "my":
				fromMyRepo = true;
				break;
			case "fs":
				isFlipswitch = true;
				break;
			}
		}
	}

	public void setOutdated() {
		available = false;
	}

	public String toString() {
		String desc = String.format("[%s] (support64: %s, available: %s, official: %s, comments: %s)", name,
				supports64Bit, available, !fromMyRepo, comments);
		desc += " { ";
		for (Entry<Integer, Support> entry : supports.entrySet())
			desc += entry.getKey() + ":" + entry.getValue().toString() + " ";
		desc += "}";
		return desc;
	}

	public void setSupport(String desc, Integer minOSVer, Integer maxOSVer) {
		assert desc != null && !desc.isEmpty();
		Matcher m;
		// example: 4:na,5:unknown,6-7:op:6.1
		// example: all:1
		boolean discountinued = false;
		boolean stopped = false;
		for (String sub : desc.split(",")) {
			if (sub.equals("dis"))
				discountinued = true;
			else if (sub.equals("stop"))
				stopped = true;
			else {
				if (sub.startsWith("all:")) {
					sub = sub.substring(4);
					String subs[] = sub.split(";");
					Support.SupportType type = Support.getType(subs[0]);
					setSupport(minOSVer, maxOSVer, type, subs.length == 2 ? subs[1] : null);
					subs = null;
				} else if ((m = iosVerPattern.matcher(sub)).find()) {
					String vers = m.group(1);
					String[] svers = vers.split("-");
					int iosVerF = Integer.parseInt(svers[0]);
					assert iosVerF >= minOSVer;
					int iosVerT = svers.length == 2 ? (svers[1].equals("l") ? maxOSVer : Integer.parseInt(svers[1])) : iosVerF;
					assert iosVerT <= maxOSVer;
					String stype = m.group(2);
					String[] stypec = stype.split(";");
					Support.SupportType type = Support.getType(stypec[0]);
					String comment = stypec.length == 2 ? stypec[1] : null;
					setSupport(iosVerF, iosVerT, type, comment);
					stypec = null;
					svers = null;
					vers = null;
					stype = null;
				}
			}
		}
		if (discountinued) {
			setOutdated();
			Support na = new Support(Support.SupportType.NA);
			for (Integer i = minOSVer; i <= maxOSVer; i++) {
				if (supports.get(i) == null)
					supports.put(i, na);
			}
		} else if (stopped) {
			Support na = new Support(Support.SupportType.NAI);
			Integer i = minOSVer;
			while (i < maxOSVer && supports.get(i) != null)
				i++;
			while (i <= maxOSVer)
				supports.put(i++, na);
		}
	}
	
	public void setSupport(String desc) {
		setSupport(desc, Version.iOSVersionMin, Version.iOSVersionMax);
	}

	public void setSupport(int iosVerF, int iosVerT, Support.SupportType type, String comment) {
		assert iosVerF <= iosVerT && iosVerF >= Version.iOSVersionMin && iosVerT <= Version.iOSVersionMax;
		Support support = new Support(type, comment);
		for (Integer i = iosVerF; i <= iosVerT; i++)
			supports.put(i, support);
	}

	public void setSupport(int iosVerF, int iosVerT, boolean support, String comment) {
		setSupport(iosVerF, iosVerT, support ? Support.SupportType.Full : Support.SupportType.No, comment);
	}

	public void setSupport(int iosVer, Support.SupportType type, String comment) {
		setSupport(iosVer, iosVer, type, comment);
	}

	public void setSupport(int iosVer, boolean support, String comment) {
		setSupport(iosVer, iosVer, support, comment);
	}

	public void setSupport(int iosVerF, int iosVerT, Support.SupportType type) {
		setSupport(iosVerF, iosVerT, type, null);
	}

	public void setSupport(int iosVerF, int iosVerT, boolean support) {
		setSupport(iosVerF, iosVerT, support, null);
	}

	public void setSupport(int iosVer, Support.SupportType type) {
		setSupport(iosVer, type, null);
	}

	public void setSupport(int iosVer, boolean support) {
		setSupport(iosVer, iosVer, support);
	}

	public void setSupport(Support.SupportType type, String comment) {
		setSupport(Version.iOSVersionMin, Version.iOSVersionMax, type, comment);
	}

	public void setSupport(boolean support, String comment) {
		setSupport(Version.iOSVersionMin, Version.iOSVersionMax, support, comment);
	}

	public void setSupport(Support.SupportType type) {
		setSupport(type, null);
	}

	public void setSupport(boolean support) {
		setSupport(support, null);
	}
}
