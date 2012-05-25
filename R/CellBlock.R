# Maybe this needs a new class, because there are conflicts with the 
# existing setCellStyle !
#


CB.setCellStyle <- function( cellBlock, style, rowIndex = NULL, colIndex = NULL )
{
    cellBlock$setCellStyle( style, .jarray( as.integer( rowIndex-1L ) ),
                           .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

CB.setFill <- function( cellBlock, fill, rowIndex = NULL, colIndex = NULL )
{
    .jcall( cellBlock, 'V', 'setFill',
            .xssfcolor( fill$foregroundColor ), .xssfcolor( fill$backgroundColor ),
            .jshort(FILL_STYLES_[[fill$pattern]]),
            .jarray( as.integer( rowIndex-1L ) ), .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

CB.setFont <- function( cellBlock, font, rowIndex = NULL, colIndex = NULL )
{
    cellBlock$setFont( font$ref, .jarray( as.integer( rowIndex-1L ) ),
                      .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

CB.putBorder <- function( cellBlock, border, rowIndex = NULL, colIndex = NULL )
{
    border_none <- BORDER_STYLES_[['BORDER_NONE']]
    borders <- c( TOP = border_none, BOTTOM = border_none,
                  LEFT = border_none, RIGHT = border_none )
    null_color <- .jnull('org/apache/poi/xssf/usermodel/XSSFColor')
    border_colors <- c( TOP = null_color, BOTTOM = null_color,
                        LEFT = null_color, RIGHT = null_color )
    borders[ border$position ] <- sapply( border$pen, function( pen ) BORDER_STYLES_[pen] )
    border_colors[ border$position ] <- sapply( border$color, .xssfcolor )
    
    .jcall( cellBlock, "V", "putBorder",
            .jshort(borders[['TOP']]), border_colors[['TOP']],
            .jshort(borders[['BOTTOM']]), border_colors[['BOTTOM']],
            .jshort(borders[['LEFT']]), border_colors[['LEFT']],
            .jshort(borders[['RIGHT']]), border_colors[['RIGHT']],
            .jarray( as.integer( rowIndex-1L ) ), .jarray( as.integer( colIndex-1L ) ) )
    invisible()
}

CB.getCell <- function( cellBlock, rowIndex, colIndex )
{
    return ( .jcall( cellBlock, 'Lorg/apache/poi/ss/usermodel/Cell;', 'getCell',
              rowIndex - 1L, colIndex - 1L ) )
}
