import org.ofdrw.converter.ConvertHelper;
import org.ofdrw.converter.FontLoader;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * OFD -> PDF。ofdrw-converter 的最小 CLI 壳；被 ofd_ops.py to-pdf 调用。
 *
 * macOS 上必须先喂默认字体：ofdrw 的 FontLoader 默认去找 simsun/simhei，
 * 找不到时 loadAsDefaultFont 收到 null 路径直接 NPE（这是好事 —— fail-closed，
 * 不像 easyofd 那样静默产出没有文字的空壳 PDF）。
 * 默认字体只是 fallback，OFD 里内嵌的字体优先。
 */
public class OfdToPdf {
    private static final String[] FALLBACK_FONTS = {
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
    };
    private static final String[] FONT_DIRS = {
        "/System/Library/Fonts", "/System/Library/Fonts/Supplemental", "/Library/Fonts",
    };

    private static void initFonts() {
        String env = System.getenv("OFD_DEFAULT_FONT");
        String picked = null;
        if (env != null && !env.isEmpty() && new File(env).isFile()) {
            picked = env;
        } else {
            for (String f : FALLBACK_FONTS) {
                if (new File(f).isFile()) { picked = f; break; }
            }
        }
        if (picked == null) {
            System.err.println("[OfdToPdf] 找不到任何可用的中文 fallback 字体 —— 拒绝转换（会产出无字空壳）");
            System.exit(3);
        }
        if (!FontLoader.loadAsDefaultFont(picked)) {
            System.err.println("[OfdToPdf] 默认字体加载失败: " + picked);
            System.exit(3);
        }
        FontLoader.setSimilarFontReplace(true);
        for (String d : FONT_DIRS) {
            File dir = new File(d);
            if (dir.isDirectory()) {
                try { FontLoader.getInstance().scanFontDir(dir); } catch (Exception ignore) {}
            }
        }
        System.err.println("[OfdToPdf] default font = " + picked);
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.err.println("usage: OfdToPdf <in.ofd> <out.pdf>");
            System.exit(2);
        }
        initFonts();
        Path in = Paths.get(args[0]);
        Path out = Paths.get(args[1]);
        ConvertHelper.toPdf(in, out);
        System.out.println("OK " + out);
    }
}
