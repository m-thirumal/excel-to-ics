/**
 * 
 */
package com.thirumal.model;

import java.time.LocalDate;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

/**
 * @author ThirumalM
 */
@Data
@Builder
@AllArgsConstructor
public class Holiday {
	
    private LocalDate date;
    private String branch;
    private String location;
    private String summary;
    private String description;
    private String imageLink;

}
