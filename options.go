package xlsx

type Options struct {
	skipLinesBeforeTitles int
	skipLinesAfterTitles int
	dateFields []string
}

type Option func(*Options)

func SkipLines(linesBeforeTitles, linesAfterTitles int) Option {
	return func(o *Options) {
		o.skipLinesBeforeTitles = linesBeforeTitles
		o.skipLinesAfterTitles = linesAfterTitles
	}
}

func SetDateFields(dateFields []string) Option {
	return func(o *Options) {
		o.dateFields = dateFields
	}
}

func getOptions(options ...Option) *Options {
	var option Options
	for _, o := range options {
		o(&option)
	}

	return &option
}

