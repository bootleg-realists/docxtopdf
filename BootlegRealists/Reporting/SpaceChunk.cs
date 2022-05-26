using iTextSharp.text.pdf.draw;

namespace iTextSharp.text;

/// <summary>
/// This class extends the chunk just by name
/// </summary>
public class SpaceChunk : Chunk
{
        /// <summary>
        /// constructors
        /// </summary>
		public SpaceChunk(): base() {}
        /// <summary>
        /// A Chunk copy constructor.
        /// </summary>
        /// <param name="ck">the Chunk to be copied</param>
		public SpaceChunk(Chunk ck): base(ck) {}
        /// <summary>
        /// Constructs a chunk of text with a certain content, without specifying a Font.
        /// </summary>
        /// <param name="content">the content</param>
		public SpaceChunk(string content): base(content) {}
        /// <summary>
        /// Constructs a chunk of text with a char, without specifying a Font .
        /// </summary>
        /// <param name="c">the content</param>
		public SpaceChunk(char c): base(c) {}
        /// <summary>
        /// Creates a separator Chunk. Note that separator chunks can't be used in combination  with tab chunks! @since 2.1.2
        /// </summary>
        /// <param name="separator">the drawInterface to use to draw the separator.</param>
		public SpaceChunk(IDrawInterface separator): base(separator) {}
        /// <summary>
        /// Creates a tab Chunk. Note that separator chunks can't be used in combination  with tab chunks! @since 2.1.2
        /// </summary>
        /// <param name="separator">the drawInterface to use to draw the tab.</param>
        /// <param name="tabPosition">an X coordinate that will be used as start position for the next Chunk.</param>
		public SpaceChunk(IDrawInterface separator, float tabPosition): base(separator, tabPosition) {}
        /// <summary>
        /// Constructs a chunk of text with a certain content and a certain Font.
        /// </summary>
        /// <param name="content">the content</param>
        /// <param name="font">the font</param>
		public SpaceChunk(string content, Font font): base(content, font) {}
        /// <summary>
        /// Constructs a chunk of text with a char and a certain Font .
        /// </summary>
        /// <param name="c">the content</param>
        /// <param name="font">the font</param>
		public SpaceChunk(char c, Font font): base(c, font) {}
        /// <summary>
        /// Creates a separator Chunk. Note that separator chunks can't be used in combination with tab chunks! @since 2.1.2
        /// </summary>
        /// <param name="separator">the drawInterface to use to draw the separator.</param>
        /// <param name="vertical">true if this is a vertical separator</param>
		public SpaceChunk(IDrawInterface separator, bool vertical): base(separator, vertical) {}
        /// <summary>
        /// Constructs a chunk containing an Image.
        /// </summary>
        /// <param name="image">the image</param>
        /// <param name="offsetX">the image offset in the x direction</param>
        /// <param name="offsetY">the image offset in the y direction</param>
		public SpaceChunk(Image image, float offsetX, float offsetY): base(image, offsetX, offsetY) {}
        /// <summary>
        /// Creates a tab Chunk. Note that separator chunks can't be used in combination with tab chunks! @since 2.1.2
        /// </summary>
        /// <param name="separator">the drawInterface to use to draw the tab.</param>
        /// <param name="tabPosition">an X coordinate that will be used as start position for the next Chunk.</param>
        /// <param name="newline"if true, a newline will be added if the tabPosition has already been reached.></param>
		public SpaceChunk(IDrawInterface separator, float tabPosition, bool newline): base(separator, tabPosition, newline) {}
        /// <summary>
        /// Constructs a chunk containing an Image.
        /// </summary>
        /// <param name="image">the image</param>
        /// <param name="offsetX">the image offset in the x direction</param>
        /// <param name="offsetY">the image offset in the y direction</param>
        /// <param name="changeLeading">true if the leading has to be adapted to the image</param>
		public SpaceChunk(Image image, float offsetX, float offsetY, bool changeLeading): base(image, offsetX, offsetY, changeLeading) {}
}