all:
	cd package && zip -r ORTools_Integration.zip .
	mv package/ORTools_Integration.zip ORTools_Integration.oxt

clean:
	rm ORTools_Integration.oxt
