package com.lframework.starter.web.inner.vo.system;

import com.lframework.starter.web.core.vo.BaseVo;
import com.lframework.starter.web.core.vo.PageVo;
import lombok.Data;

import java.io.Serializable;

@Data
public class QuerySelectorVo  extends PageVo implements BaseVo, Serializable {

  private String label;

}
