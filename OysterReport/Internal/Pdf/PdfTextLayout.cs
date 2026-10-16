namespace OysterReport.Internal.Pdf;

internal enum TextHorizontal
{
    Near,
    Center,
    Far
}

internal enum TextVertical
{
    Near,
    Center,
    Far
}

internal enum TextParagraph
{
    Left,
    Center,
    Right,
    Justify
}

internal static class PdfTextLayout
{
    public static void DrawString(PdfPageCanvas canvas, string text, PdfFont font, PdfColor color, double x, double y, double width, double height, TextHorizontal horizontal, TextVertical vertical)
    {
        var textWidth = font.MeasureWidth(text);
        switch (horizontal)
        {
            case TextHorizontal.Center:
                x += (width - textWidth) / 2;
                break;
            case TextHorizontal.Far:
                x += width - textWidth;
                break;
        }

        switch (vertical)
        {
            case TextVertical.Near:
                y += font.CellAscent;
                break;
            case TextVertical.Center:
                y += (font.CellAscent * 3 / 4 / 2) + (height / 2);
                break;
            case TextVertical.Far:
                y += -font.CellDescent + height;
                break;
        }

        Show(canvas, text, font, color, x, y);
    }

    public static void Show(PdfPageCanvas canvas, string text, PdfFont font, PdfColor color, double x, double baseline) =>
        canvas.ShowText(font.Face, font.Size, color, x, baseline, text, font.SimulateItalic ? PdfFont.ItalicSkew : 0);

    public static void DrawWrapped(PdfPageCanvas canvas, string text, PdfFont font, PdfColor color, double x, double y, double width, double height, TextParagraph alignment)
    {
        if (text.Length == 0)
        {
            return;
        }

        var spaceWidth = font.MeasureWidth("x x") - font.MeasureWidth("xx");
        var blocks = CreateBlocks(text, font);
        CreateLayout(blocks, font, width, height, spaceWidth, alignment);

        var dy = y + font.CellAscent;
        foreach (var block in blocks)
        {
            if (block.Stop)
            {
                break;
            }

            if (block.IsLineBreak)
            {
                continue;
            }

            Show(canvas, block.Text, font, color, x + block.X, dy + block.Y);
        }
    }

    private static List<Block> CreateBlocks(string text, PdfFont font)
    {
        var blocks = new List<Block>();
        var inNonWhiteSpace = false;
        int startIndex = 0, blockLength = 0;
        for (var idx = 0; idx < text.Length; idx++)
        {
            var ch = text[idx];
            if (ch == '\r')
            {
                if (idx < text.Length - 1 && text[idx + 1] == '\n')
                {
                    idx++;
                }

                ch = '\n';
            }

            if (ch == '\n')
            {
                if (blockLength != 0)
                {
                    var token = text.Substring(startIndex, blockLength);
                    blocks.Add(new Block(token, font.MeasureWidth(token)));
                }

                startIndex = idx + 1;
                blockLength = 0;
                blocks.Add(Block.LineBreak());
            }
            else if (ch != '\u00A0' && Char.IsWhiteSpace(ch))
            {
                if (inNonWhiteSpace)
                {
                    var token = text.Substring(startIndex, blockLength);
                    blocks.Add(new Block(token, font.MeasureWidth(token)));
                    startIndex = idx + 1;
                    blockLength = 0;
                }
                else
                {
                    blockLength++;
                }
            }
            else
            {
                inNonWhiteSpace = true;
                blockLength++;
            }
        }

        if (blockLength != 0)
        {
            var token = text.Substring(startIndex, blockLength);
            blocks.Add(new Block(token, font.MeasureWidth(token)));
        }

        return blocks;
    }

    private static void CreateLayout(List<Block> blocks, PdfFont font, double rectWidth, double layoutHeight, double spaceWidth, TextParagraph alignment)
    {
        var rectHeight = layoutHeight - font.CellAscent - font.CellDescent;
        var lineSpace = font.LineSpace;
        var firstIndex = 0;
        double x = 0, y = 0;
        for (var idx = 0; idx < blocks.Count; idx++)
        {
            var block = blocks[idx];
            if (block.IsLineBreak)
            {
                if (alignment == TextParagraph.Justify)
                {
                    blocks[firstIndex].ForceLeft = true;
                }

                AlignLine(blocks, firstIndex, idx - 1, rectWidth, spaceWidth, alignment);
                firstIndex = idx + 1;
                x = 0;
                y += lineSpace;
                if (y > rectHeight)
                {
                    block.Stop = true;
                    break;
                }
            }
            else
            {
                var width = block.Width;
                if (x + width <= rectWidth || x == 0)
                {
                    block.X = x;
                    block.Y = y;
                    x += width + spaceWidth;
                }
                else
                {
                    AlignLine(blocks, firstIndex, idx - 1, rectWidth, spaceWidth, alignment);
                    firstIndex = idx;
                    y += lineSpace;
                    if (y > rectHeight)
                    {
                        block.Stop = true;
                        break;
                    }

                    block.X = 0;
                    block.Y = y;
                    x = width + spaceWidth;
                }
            }
        }

        if (firstIndex < blocks.Count && alignment != TextParagraph.Justify)
        {
            AlignLine(blocks, firstIndex, blocks.Count - 1, rectWidth, spaceWidth, alignment);
        }
    }

    private static void AlignLine(List<Block> blocks, int firstIndex, int lastIndex, double layoutWidth, double spaceWidth, TextParagraph alignment)
    {
        if (firstIndex >= blocks.Count)
        {
            return;
        }

        if (alignment == TextParagraph.Left || blocks[firstIndex].ForceLeft)
        {
            return;
        }

        var count = lastIndex - firstIndex + 1;
        if (count <= 0)
        {
            return;
        }

        var totalWidth = -spaceWidth;
        for (var idx = firstIndex; idx <= lastIndex; idx++)
        {
            totalWidth += blocks[idx].Width + spaceWidth;
        }

        var dx = Math.Max(layoutWidth - totalWidth, 0);
        if (alignment != TextParagraph.Justify)
        {
            if (alignment == TextParagraph.Center)
            {
                dx /= 2;
            }

            for (var idx = firstIndex; idx <= lastIndex; idx++)
            {
                blocks[idx].X += dx;
            }
        }
        else if (count > 1)
        {
            dx /= count - 1;
            for (int idx = firstIndex + 1, i = 1; idx <= lastIndex; idx++, i++)
            {
                blocks[idx].X += dx * i;
            }
        }
    }

    private sealed class Block(string text, double width)
    {
        public string Text { get; } = text;

        public double Width { get; } = width;

        public bool IsLineBreak { get; private init; }

        public double X { get; set; }

        public double Y { get; set; }

        public bool Stop { get; set; }

        public bool ForceLeft { get; set; }

        public static Block LineBreak() => new(string.Empty, 0) { IsLineBreak = true };
    }
}
