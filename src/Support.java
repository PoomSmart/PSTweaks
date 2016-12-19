
public class Support {
	public enum SupportType {
		Full, CodePartial, OSPartial, Maybe, Unknown, No, NAI, NA
	};

	public SupportType type;
	public String comment;

	public Support(SupportType type, String comment) {
		this.type = type == null ? SupportType.NA : type;
		this.comment = comment == null && isPartial(type) ? "Partial" : comment;
	}

	public Support(SupportType type) {
		this(type, null);
	}

	public static boolean isPartial(SupportType type) {
		return type == SupportType.CodePartial || type == SupportType.OSPartial;
	}

	public static SupportType getType(String s) {
		switch (s) {
		case "1":
			return SupportType.Full;
		case "cp":
			return SupportType.CodePartial;
		case "op":
			return SupportType.OSPartial;
		case "m":
			return SupportType.Maybe;
		case "?":
			return SupportType.Unknown;
		case "0":
			return SupportType.No;
		case "na":
			return SupportType.NA;
		default:
			return null;
		}
	}

	public String toString() {
		switch (type) {
		case Full:
			return "Yes";
		case CodePartial:
		case OSPartial:
			return comment;
		case Maybe:
			return "Maybe";
		case Unknown:
			return "?";
		case No:
			return "No";
		case NAI:
			return "N/Ai";
		case NA:
			return "N/A";
		default:
			return "UNK";
		}
	}
}
