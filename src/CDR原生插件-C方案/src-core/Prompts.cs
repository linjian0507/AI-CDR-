using System;
using System.Globalization;
using System.Text;

namespace AIVectorCore
{
    public static class PromptBuilder
    {
        public static string BuildSystemPrompt(GenerationOptions options)
        {
            options = options ?? new GenerationOptions();
            var width = options.Width > 0 ? options.Width : 1024;
            var height = options.Height > 0 ? options.Height : 1024;
            var layerCount = string.IsNullOrWhiteSpace(options.LayerCount) ? "5" : options.LayerCount;
            var isCopyMode = string.Equals(options.ReferenceMode, "copy", StringComparison.OrdinalIgnoreCase);
            var lines = new[]
            {
                "你是一位顶级矢量插画师, 任务是输出可直接被 CorelDRAW 导入的、严格合法的 SVG 代码。硬性要求:",
                "1. 只输出一个完整的 <svg>…</svg> 代码块, 前后不要有任何解释文字。",
                string.Format(CultureInfo.InvariantCulture, "2. 根元素必须为: <svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {0} {1}\">, 不要设置 width/height。", width, height),
                "3. 画面必须分层: 每个图层是 <svg> 的一个直接子元素 <g>, 按绘制顺序从底层到顶层排列(第一个 <g> 是最底层)。",
                "4. 图层数量约为 " + layerCount + " 层, 典型划分: 背景 / 远景 / 主体 / 细节 / 高光装饰 / 文字。",
                "5. 每个 <g> 必须带 id(英文小写)和 data-name(简短中文图层名, 例如 data-name=\"背景\")。",
                "6. 只允许 path、rect、circle、ellipse、polygon、polyline、line、text 元素; 渐变可用 <defs> 中的 linearGradient/radialGradient。",
                "7. 禁止使用: <image>、<style>、CSS class、filter 滤镜、mask、clipPath、外部引用、<script>、动画。所有颜色用 fill/stroke 内联属性表示。",
                "8. 造型精致、路径圆滑、细节丰富, 配色专业和谐; 避免元素超出 viewBox。",
                "9. 除 <defs> 外, 不要在 <svg> 直接子级放置任何非 <g> 的图形元素。"
            };

            var result = new StringBuilder(string.Join("\n", lines));
            if (!string.IsNullOrWhiteSpace(options.StyleDescription))
                result.Append("\n10. 整体艺术风格: ").Append(options.StyleDescription).Append("。");
            if (!string.IsNullOrWhiteSpace(options.Palette))
                result.Append("\n11. 配色要求: ").Append(options.Palette).Append("。");
            if (options.NoBackground)
                result.Append("\n12. 不要绘制背景层, 保持背景透明, 只画主体内容。");

            if (isCopyMode)
            {
                result.Append("\n\n【临摹还原模式追加硬性要求】")
                    .Append("\nA. 当前是视觉描摹/版式复刻任务, 不是重新设计、不是概念再创作。参考图是唯一目标, 必须优先还原它的画布比例、留白、边距、中心轴、层级遮挡、位置和尺寸关系。")
                    .Append("\nB. 先建立整体版式骨架, 再画局部细节: 先估算每个大对象的边界框(x/y/宽/高), 大背景色块、标题牌/卡片/标签、装饰圆点和植物、人物/道具/图标、底部文字、纸张边框都要按原图相对坐标摆放, 关键对象边界框偏差尽量控制在 10% 内。")
                    .Append("\nC. 所有可识别文字必须使用 <text> 元素保持可编辑; 先内部 OCR 出每一行原文, 再逐字照抄参考图, 不要改写、翻译、增删或编造。中文直接写中文, 用 font-size、font-weight、letter-spacing、text-anchor 接近原图粗细和位置; 不要把文字转成路径。")
                    .Append("\nD. 人物和主体物不能套用通用模板: 数量、性别气质、姿势、发型、服装颜色、器械/图标结构、遮挡关系都以参考图为准。不要添加参考图没有的口罩、听诊器、领带、图标、光斑或装饰; 难以精描时可简化, 但不能删除或换成其它元素。")
                    .Append("\nE. 颜色必须从参考图取近似色; 大色块、文字颜色、人物服饰和装饰形状的主色优先准确, 不要额外添加参考图没有的强阴影、3D、透视、滤镜或纹理。")
                    .Append("\nF. SVG 要像可生产的 CorelDRAW 矢量稿: 路径尽量闭合平滑, 图形可拆分编辑; 同类元素分在清晰图层中, 但不要为了层数强行合并丢失细节。")
                    .Append("\nG. 海报、卡片、展板类图片要保留纸张/画板边界、顶部主标题区、中心插图区、底部署名区的纵向比例; 不要为了画得更满而放大主体或挤掉留白。")
                    .Append("\nH. 输出前自检: 文字是否可编辑且原文一致; 大标题/主体/人物/底部信息是否齐全; 关键位置比例是否接近参考图; 是否存在乱码、占位文字或无关元素。");
            }

            return result.ToString();
        }

        public static string BuildUserPrompt(string prompt, string referenceMode, bool hasReference)
        {
            prompt = (prompt ?? "").Trim();
            if (hasReference)
            {
                var lead = string.Equals(referenceMode, "copy", StringComparison.OrdinalIgnoreCase)
                    ? "请把参考图片作为唯一目标, 执行【视觉描摹/版式复刻】并输出可编辑分层 SVG。不要按主题重新设计, 不要换构图。请先内部识别: 画布/纸张边框、每个大对象边界框、所有文字行、人物和道具, 再按这些位置生成 SVG。必须尽量复刻: ①整体画布比例、边距、背景色块和装饰图形; ②标题牌/卡片/挂图/图标等几何结构的位置和尺寸; ③人物数量、站位、姿势、发型、服装颜色和道具, 不添加参考图没有的口罩/器械/装饰; ④所有可识别文字, 必须用 <text> 逐字重现, 中文保持中文且可编辑。无法精确识别的细节用简洁矢量近似表达, 但不要删减主要元素。"
                    : "请参考这张图片的风格、配色与构图气质, 创作一幅矢量 SVG 插画, 并按要求分层。";
                return lead + (prompt.Length == 0 ? "" : "\n补充要求: " + prompt);
            }
            return prompt.Length == 0 ? "" : "绘制内容: " + prompt;
        }
        public static string BuildEditPrompt(string svg, string instruction)
        {
            return "现有 SVG 代码:\n```svg\n" + (svg ?? "") +
                   "\n```\n\n修改要求: " + (instruction ?? "").Trim() +
                   "\n\n请输出修改后的完整 SVG。";
        }
    }
}
