import org.ofdrw.converter.ImageMaker;
import org.ofdrw.converter.FontLoader;
import org.ofdrw.reader.OFDReader;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

/** OFD -> 每页 PNG。用于 to-pdf 的图片路线（ofdrw 自绘，绕开 PDF 文本定位漂移）。 */
public class OfdToImages {
    private static void initFonts() {
        String env = System.getenv("OFD_DEFAULT_FONT");
        String[] cands = {env, "/Library/Fonts/Arial Unicode.ttf"};
        for (String f : cands) {
            if (f != null && !f.isEmpty() && new File(f).isFile() && FontLoader.loadAsDefaultFont(f)) {
                System.err.println("[OfdToImages] default font = " + f);
                return;
            }
        }
        System.err.println("[OfdToImages] 无可用 fallback 字体，拒绝渲染");
        System.exit(3);
    }

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("usage: OfdToImages <in.ofd> <outDir> [ppm=8 (~200dpi)]");
            System.exit(2);
        }
        initFonts();
        // ImageMaker(reader,int) 的第二参是 **ppm（像素/毫米）不是 dpi** —— 传 150 会得到
        // 44556x31485 的巨图并 OOM。200dpi ≈ 7.87 px/mm，取 8。
        int ppm = args.length > 2 ? Integer.parseInt(args[2]) : 8;
        Path in = Paths.get(args[0]);
        File outDir = new File(args[1]);
        outDir.mkdirs();
        try (OFDReader reader = new OFDReader(in)) {
            ImageMaker maker = new ImageMaker(reader, ppm);
            int n = reader.getNumberOfPages();
            for (int i = 0; i < n; i++) {
                BufferedImage img = maker.makePage(i);
                File f = new File(outDir, String.format("page-%03d.png", i + 1));
                ImageIO.write(img, "PNG", f);
            }
            System.out.println("OK " + n + " pages -> " + outDir);
        }
    }
}
