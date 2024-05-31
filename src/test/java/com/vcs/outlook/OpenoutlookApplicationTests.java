package com.vcs.outlook;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.opentest4j.AssertionFailedError;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import com.vcs.outlook.controller.InstrumentDetails;
import com.vcs.outlook.controller.WriteDataToExcel;

@SpringBootTest
class OpenoutlookApplicationTests {
	
	@Autowired
	WriteDataToExcel dataToExcel;

	@Test
	void contextLoads() throws Exception {
		

        List<InstrumentDetails> instrumentList = new ArrayList<>();
        instrumentList.add(new InstrumentDetails(5027, "Guitar", "Vishnu", 299.99, LocalDate.of(2024, 1, 1)));
        instrumentList.add(new InstrumentDetails(5028, "Piano", "lal", 499.99, LocalDate.of(2024, 2, 1)));
        instrumentList.add(new InstrumentDetails(5029, "Flute", "Wind", 149.99, LocalDate.of(2024, 3, 1)));

        AssertionFailedError exception = Assertions.assertThrows(AssertionFailedError.class, () -> {
            Assertions.assertEquals(1, 2, "message");
        });
        
        String string = exception.getMessage();
		dataToExcel.writeToW("sheeNew", instrumentList);

	}

}
