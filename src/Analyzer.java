import java.util.TreeMap;

public class Analyzer {
	public TreeMap<String, Tweak> tweaks;

	public Analyzer(TreeMap<String, Tweak> tweaks) {
		this.tweaks = tweaks;
	}

	public boolean isApplicableType(Support.SupportType type, boolean guess) {
		return type == Support.SupportType.Full || type == Support.SupportType.CodePartial
				|| type == Support.SupportType.OSPartial || (guess && type == Support.SupportType.Maybe);
	}

	public void printTweaksSupportiOSRange(int iOSVerF, int iOSVerT, boolean guess) {
		assert iOSVerF <= iOSVerT && iOSVerF >= Version.iOSVersionMin && iOSVerT <= Version.iOSVersionMax;
		int total = 0;
		for (Tweak tweak : tweaks.values()) {
			boolean pass = true;
			for (Integer i = iOSVerF; i <= iOSVerT; i++) {
				Support support = tweak.supports.get(i);
				if (support == null) {
					System.out.println("Error: Support for iOS version " + i + " not found for tweak " + tweak.name);
					continue;
				}
				if (!isApplicableType(support.type, guess)) {
					pass = false;
					break;
				}
			}
			if (!pass)
				continue;
			System.out.println(tweak);
			total++;
		}
		System.out.println("Total tweaks: " + total);
	}
}
