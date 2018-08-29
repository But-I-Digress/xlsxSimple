
getColSpec <- function (n) {
	if (n < 27) return(LETTERS[[n]])
	
	l <- LETTERS
	for (i in 1:26) l <- c(l, paste(LETTERS[[i]], LETTERS, sep=""))
	if (n < 703) return(l[[n]])
}
