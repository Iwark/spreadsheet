package spreadsheet

// DimensionProperties is properties about a dimension.
type DimensionProperties struct {
	HiddenByFilter bool `json:"hiddenByFilter"`
	HiddenByUser   bool `json:"hiddenByUser"`
	PixelSize      uint `json:"pixelSize"`
	// DeveloperMetadata []*DeveloperMetadata `json:"developerMetadata"`
}
